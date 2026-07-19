import time
import traceback

from fastapi import FastAPI, Request, HTTPException
import requests
from openai import OpenAI
from fastapi.responses import Response, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import os
import json
from dotenv import load_dotenv
import re
import unicodedata

from fastapi.responses import HTMLResponse

from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from io import BytesIO

from rapidfuzz import process, fuzz

from fastapi.staticfiles import StaticFiles

load_dotenv()

ZONA_HORARIA = ZoneInfo("America/Bogota")

def hora_local():
    return datetime.now(ZONA_HORARIA).strftime("%Y-%m-%d %H:%M:%S")

# with open("menu.json", "r", encoding="utf-8") as file:
#     MENU = json.load(file)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    timeout=25.0,
    max_retries=1
)
DATABASE_URL = os.getenv("DATABASE_URL")
META_ACCESS_TOKEN = os.getenv("META_ACCESS_TOKEN")
META_PHONE_NUMBER_ID = os.getenv("META_PHONE_NUMBER_ID")
META_VERIFY_TOKEN = os.getenv("META_VERIFY_TOKEN")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://consult-unsent-undying.ngrok-free.dev")

def get_connection():
    return psycopg2.connect(DATABASE_URL)

conversaciones = {}

CAMPOS_OBLIGATORIOS = [
    "nombre",
    "items",
    "tipo_entrega",
    "direccion"
]

def crear_tabla():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS orders (
            id SERIAL PRIMARY KEY,
            fecha TEXT,
            telefono TEXT,
            nombre TEXT,
            tipo_entrega TEXT,
            direccion TEXT,
            hora TEXT,
            metodo_pago TEXT,
            notas TEXT,
            estado TEXT,
            total NUMERIC DEFAULT 0
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS order_items (
            id SERIAL PRIMARY KEY,
            order_id INTEGER REFERENCES orders(id),
            producto TEXT,
            cantidad TEXT,
            notas TEXT
        )
    """)

    cursor.execute("""
        ALTER TABLE order_items
        ADD COLUMN IF NOT EXISTS notas TEXT
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS conversation_states (
            telefono TEXT PRIMARY KEY,
            pedido_json JSONB,
            updated_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id SERIAL PRIMARY KEY,
            telefono TEXT,
            nombre TEXT,
            mensaje TEXT,
            direccion TEXT,
            fecha TEXT
        )
    """)

    cursor.execute("""
        ALTER TABLE messages
        ADD COLUMN IF NOT EXISTS leido BOOLEAN DEFAULT FALSE
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            telefono TEXT PRIMARY KEY,
            nombre TEXT,
            direccion TEXT,
            creado_en TEXT,
            actualizado_en TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS whatsapp_accounts (
            id SERIAL PRIMARY KEY,
            waba_id TEXT UNIQUE NOT NULL,
            phone_number_id TEXT,
            business_id TEXT,
            access_token TEXT NOT NULL,
            token_expires_at TIMESTAMP,
            connected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            active BOOLEAN DEFAULT TRUE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_whatsapp_messages (
            message_id TEXT PRIMARY KEY,
            telefono TEXT,
            processed_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS human_handoffs (
            telefono TEXT PRIMARY KEY,
            activo_hasta TIMESTAMP,
            activado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            motivo TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id SERIAL PRIMARY KEY,
            producto TEXT UNIQUE NOT NULL,
            precio INTEGER NOT NULL CHECK (precio >= 0),
            activo BOOLEAN NOT NULL DEFAULT TRUE,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()

def importar_menu_inicial():
    """
    Importa los productos de menu.json solamente si todavía
    no existen en la tabla products.
    """

    ruta_menu = "menu.json"

    if not os.path.exists(ruta_menu):
        print("No se encontró menu.json para la importación inicial.")
        return

    try:
        with open(ruta_menu, "r", encoding="utf-8") as file:
            productos_json = json.load(file)
    except (OSError, json.JSONDecodeError) as error:
        print("ERROR LEYENDO menu.json:", error)
        return

    conn = get_connection()
    cursor = conn.cursor()

    for item in productos_json:
        producto = str(item.get("producto", "")).strip()

        try:
            precio = int(item.get("precio", 0))
        except (TypeError, ValueError):
            continue

        if not producto or precio < 0:
            continue

        cursor.execute("""
            INSERT INTO products (
                producto,
                precio,
                activo
            )
            VALUES (%s, %s, TRUE)
            ON CONFLICT (producto) DO NOTHING
        """, (
            producto,
            precio
        ))

    conn.commit()
    conn.close()

    print("Menú inicial importado correctamente.")

crear_tabla()
importar_menu_inicial()

def obtener_conversacion(numero):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute(
        "SELECT pedido_json FROM conversation_states WHERE telefono = %s",
        (numero,)
    )

    row = cursor.fetchone()
    conn.close()

    if row:
        return row["pedido_json"]

    return {
        "nombre": "",
        "items": [],
        "productos_ambiguos": [],
        "productos_no_disponibles": [],
        "tipo_entrega": "",
        "direccion": "",
        "hora": "",
        "metodo_pago": "",
        "notas": "",
        "esperando_confirmacion": False,
        "pedido_confirmado": False,
        "esperando_metodo_pago": False,
        "fecha_confirmacion": ""
    }

def guardar_mensaje(telefono, nombre, mensaje, direccion):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO messages (telefono, nombre, mensaje, direccion, fecha, leido)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (
        telefono,
        nombre,
        mensaje,
        direccion,
        hora_local(),
        True if direccion == "out" else False
    ))

    conn.commit()
    conn.close()

def obtener_contacto(telefono):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT *
        FROM contacts
        WHERE telefono = %s
    """, (telefono,))

    contacto = cursor.fetchone()
    conn.close()

    return contacto

def obtener_ultimo_pedido_reciente(numero, minutos=15):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT *
        FROM orders
        WHERE telefono = %s
          AND fecha::timestamp >= NOW() - (%s * INTERVAL '1 minute')
        ORDER BY id DESC
        LIMIT 1
    """, (numero, minutos))

    pedido = cursor.fetchone()
    conn.close()

    return pedido


def guardar_contacto(telefono, nombre=None, direccion=None):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO contacts (telefono, nombre, direccion, creado_en, actualizado_en)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (telefono)
        DO UPDATE SET
            nombre = COALESCE(EXCLUDED.nombre, contacts.nombre),
            direccion = COALESCE(EXCLUDED.direccion, contacts.direccion),
            actualizado_en = EXCLUDED.actualizado_en
    """, (
        telefono,
        nombre,
        direccion,
        hora_local(),
        hora_local()
    ))

    conn.commit()
    conn.close()


def guardar_conversacion(numero, pedido):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO conversation_states (telefono, pedido_json, updated_at)
        VALUES (%s, %s, %s)
        ON CONFLICT (telefono)
        DO UPDATE SET
            pedido_json = EXCLUDED.pedido_json,
            updated_at = EXCLUDED.updated_at
    """, (
        numero,
        json.dumps(pedido, ensure_ascii=False),
        hora_local()
    ))

    conn.commit()
    conn.close()


def borrar_conversacion(numero):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM conversation_states WHERE telefono = %s",
        (numero,)
    )

    conn.commit()
    conn.close()

def activar_atencion_humana(telefono, minutos=30, motivo="Solicitud del cliente"):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO human_handoffs (
            telefono,
            activo_hasta,
            activado_en,
            motivo
        )
        VALUES (
            %s,
            NOW() + (%s * INTERVAL '1 minute'),
            NOW(),
            %s
        )
        ON CONFLICT (telefono)
        DO UPDATE SET
            activo_hasta = NOW() + (%s * INTERVAL '1 minute'),
            activado_en = NOW(),
            motivo = EXCLUDED.motivo
    """, (
        telefono,
        minutos,
        motivo,
        minutos
    ))

    conn.commit()
    conn.close()


def desactivar_atencion_humana(telefono):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        DELETE FROM human_handoffs
        WHERE telefono = %s
    """, (telefono,))

    conn.commit()
    conn.close()


def obtener_estado_atencion_humana(telefono):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT
            telefono,
            activo_hasta,
            activado_en,
            motivo,
            activo_hasta > NOW() AS activo
        FROM human_handoffs
        WHERE telefono = %s
    """, (telefono,))

    estado = cursor.fetchone()
    conn.close()

    if not estado:
        return {
            "activo": False,
            "activo_hasta": None,
            "motivo": ""
        }

    return estado


def atencion_humana_activa(telefono):
    estado = obtener_estado_atencion_humana(telefono)
    return bool(estado and estado.get("activo"))


def renovar_atencion_humana(telefono, minutos=30):
    """
    Renueva el periodo solamente cuando el chat ya está
    bajo atención humana.
    """
    if atencion_humana_activa(telefono):
        activar_atencion_humana(
            telefono,
            minutos=minutos,
            motivo="Conversación atendida manualmente"
        )

def mensaje_fue_procesado(message_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 1
        FROM processed_whatsapp_messages
        WHERE message_id = %s
    """, (message_id,))

    existe = cursor.fetchone() is not None

    conn.close()

    return existe


def marcar_mensaje_procesado(message_id, telefono=None):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO processed_whatsapp_messages (
            message_id,
            telefono,
            processed_at
        )
        VALUES (%s, %s, %s)
        ON CONFLICT (message_id) DO NOTHING
    """, (
        message_id,
        telefono,
        hora_local()
    ))

    conn.commit()
    conn.close()

def finalizar_webhook(message_id, numero):
    if message_id:
        marcar_mensaje_procesado(message_id, numero)

    return Response(
        content="EVENT_RECEIVED",
        media_type="text/plain"
    )

def normalizar(texto):
    texto = texto.lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"\b(un|una|unos|unas|porfa|por favor)\b", "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto

templates = Jinja2Templates(directory="templates")

def calcular_total(pedido):
    total = 0
    menu_activo = obtener_menu(solo_activos=True)

    for item in pedido.get("items", []):
        producto = item.get("producto", "")

        try:
            cantidad = int(item.get("cantidad", 1))
        except (TypeError, ValueError):
            cantidad = 1

        producto_encontrado = None

        for producto_menu in menu_activo:
            if normalizar(producto_menu["producto"]) == normalizar(producto):
                producto_encontrado = producto_menu
                break

        precio = producto_encontrado["precio"] if producto_encontrado else 0

        extra = 0
        notas_item = item.get("notas", "").lower()

        if "deslactosada" in notas_item:
            extra = 2000

        total += (precio + extra) * cantidad

    return total

def buscar_producto(nombre_producto):
    menu_activo = obtener_menu(solo_activos=True)
    nombre_producto = normalizar(nombre_producto)

    for item in menu_activo:
        if normalizar(item["producto"]) == nombre_producto:
            return item

    nombres_normalizados = [
        normalizar(item["producto"])
        for item in menu_activo
    ]

    if not nombres_normalizados:
        return None

    resultado = process.extractOne(
        nombre_producto,
        nombres_normalizados,
        scorer=fuzz.WRatio
    )

    if resultado is None:
        return None

    _, score, index = resultado

    if score >= 75:
        return menu_activo[index]

    return None

def validar_items_pedido(pedido):
    menu_activo = obtener_menu(solo_activos=True)

    items_validos = []
    productos_invalidos = []

    for item in pedido.get("items", []):
        nombre_producto = item.get("producto", "").strip()

        producto_encontrado = None

        for producto_menu in menu_activo:
            if normalizar(producto_menu["producto"]) == normalizar(nombre_producto):
                producto_encontrado = producto_menu
                break

        if producto_encontrado:
            items_validos.append({
                "producto": producto_encontrado["producto"],
                "cantidad": str(item.get("cantidad", "1")),
                "notas": item.get("notas", "")
            })
        else:
            productos_invalidos.append(nombre_producto)

    pedido["items"] = items_validos

    productos_no_disponibles = pedido.get(
        "productos_no_disponibles",
        []
    )

    for producto in productos_invalidos:
        if producto and producto not in productos_no_disponibles:
            productos_no_disponibles.append(producto)

    pedido["productos_no_disponibles"] = productos_no_disponibles

    return pedido

def opciones_por_producto(producto_ambiguo):
    menu_activo = obtener_menu(solo_activos=True)
    producto_ambiguo = normalizar(producto_ambiguo)

    opciones = []

    for item in menu_activo:
        nombre = item["producto"]
        nombre_normalizado = normalizar(nombre)

        if producto_ambiguo in nombre_normalizado:
            opciones.append(nombre)

    return opciones

def guardar_pedido_db(numero, pedido):
    conn = get_connection()
    cursor = conn.cursor()
    total = calcular_total(pedido)

    cursor.execute("""
        INSERT INTO orders (
            fecha, telefono, nombre, tipo_entrega,
            direccion, hora, metodo_pago, notas, estado, total
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id
    """, (
        hora_local(),
        numero,
        pedido.get("nombre", ""),
        pedido.get("tipo_entrega", ""),
        pedido.get("direccion", ""),
        pedido.get("hora", ""),
        pedido.get("metodo_pago", ""),
        pedido.get("notas", ""),
        "recibido",
        total
    ))

    order_id = cursor.fetchone()[0]

    for item in pedido.get("items", []):
        cursor.execute("""
            INSERT INTO order_items (
                order_id, producto, cantidad, notas
            ) VALUES (%s, %s, %s, %s)
        """, (
            order_id,
            item.get("producto", ""),
            item.get("cantidad", ""),
            item.get("notas", "")
        ))

    conn.commit()
    conn.close()

# Funciones para META

def enviar_texto(numero, texto):
    url = f"https://graph.facebook.com/v25.0/{META_PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "text",
        "text": {
            "body": texto
        }
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=30
        )

        print("META RESPONSE:", response.status_code, response.text)

        if response.status_code in [200, 201]:
            guardar_mensaje(numero, "Terrace", texto, "out")
            return True

        return False

    except requests.RequestException as error:
        print("ERROR ENVIANDO TEXTO:", str(error))
        return False
    
def obtener_menu(solo_activos=True):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    if solo_activos:
        cursor.execute("""
            SELECT id, producto, precio, activo
            FROM products
            WHERE activo = TRUE
            ORDER BY producto ASC
        """)
    else:
        cursor.execute("""
            SELECT id, producto, precio, activo
            FROM products
            ORDER BY producto ASC
        """)

    productos = cursor.fetchall()
    conn.close()

    return [dict(producto) for producto in productos]


def obtener_producto_por_id(product_id):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT id, producto, precio, activo
        FROM products
        WHERE id = %s
    """, (product_id,))

    producto = cursor.fetchone()
    conn.close()

    return dict(producto) if producto else None

def enviar_menu(numero, link=None):
    ruta_imagen = os.path.join("static", "menu.jpeg")

    if not os.path.exists(ruta_imagen):
        print("ERROR: No existe el menú:", ruta_imagen)
        return False

    upload_url = (
        f"https://graph.facebook.com/v25.0/"
        f"{META_PHONE_NUMBER_ID}/media"
    )

    headers_auth = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}"
    }

    try:
        # 1. Subir el menu a Meta
        with open(ruta_imagen, "rb") as archivo:
            files = {
                "file": (
                    "menu.jpeg",
                    archivo,
                    "image/jpeg"
                )
            }

            data = {
                "messaging_product": "whatsapp",
                "type": "image/jpeg"
            }

            upload_response = requests.post(
                upload_url,
                headers=headers_auth,
                files=files,
                data=data,
                timeout=60
            )

        print(
            "META MENU UPLOAD:",
            upload_response.status_code,
            upload_response.text
        )

        if upload_response.status_code not in [200, 201]:
            print("ERROR SUBIENDO MENÚ A META")
            return False

        media_id = upload_response.json().get("id")

        if not media_id:
            print("ERROR: Meta no devolvió media_id para el menú")
            return False

        # 2. Enviar el menú usando el media_id
        messages_url = (
            f"https://graph.facebook.com/v25.0/"
            f"{META_PHONE_NUMBER_ID}/messages"
        )

        headers_json = {
            "Authorization": f"Bearer {META_ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": numero,
            "type": "image",
            "image": {
                "id": media_id,
                "caption": "Menú de Terrace"
            }
        }

        send_response = requests.post(
            messages_url,
            headers=headers_json,
            json=payload,
            timeout=30
        )

        print(
            "META MENU SEND:",
            send_response.status_code,
            send_response.text
        )

        if send_response.status_code in [200, 201]:
            guardar_mensaje(
                numero,
                "Terrace",
                "🧾 Menú de Terrace enviado",
                "out"
            )
            return True

        print("ERROR ENVIANDO MENÚ POR MEDIA_ID")
        return False

    except requests.RequestException as error:
        print("ERROR DE CONEXIÓN ENVIANDO MENÚ:", str(error))
        return False

    except Exception as error:
        print("ERROR GENERAL ENVIANDO MENÚ:", str(error))
        return False
    
def enviar_imagen(numero, link=None):
    ruta_imagen = os.path.join("static", "qr_transferencia.jpeg")

    if not os.path.exists(ruta_imagen):
        print("ERROR: No existe el QR:", ruta_imagen)
        return False

    upload_url = (
        f"https://graph.facebook.com/v25.0/"
        f"{META_PHONE_NUMBER_ID}/media"
    )

    headers_auth = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}"
    }

    try:
        # 1. Subir el QR a Meta
        with open(ruta_imagen, "rb") as archivo:
            files = {
                "file": (
                    "qr_transferencia.jpeg",
                    archivo,
                    "image/jpeg"
                )
            }

            data = {
                "messaging_product": "whatsapp",
                "type": "image/jpeg"
            }

            upload_response = requests.post(
                upload_url,
                headers=headers_auth,
                files=files,
                data=data,
                timeout=60
            )

        print(
            "META QR UPLOAD:",
            upload_response.status_code,
            upload_response.text
        )

        if upload_response.status_code not in [200, 201]:
            print("ERROR SUBIENDO QR A META")
            return False

        media_id = upload_response.json().get("id")

        if not media_id:
            print("ERROR: Meta no devolvió media_id para el QR")
            return False

        # 2. Enviar el QR usando el media_id
        messages_url = (
            f"https://graph.facebook.com/v25.0/"
            f"{META_PHONE_NUMBER_ID}/messages"
        )

        headers_json = {
            "Authorization": f"Bearer {META_ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": numero,
            "type": "image",
            "image": {
                "id": media_id,
                "caption": "Código QR para transferencia"
            }
        }

        send_response = requests.post(
            messages_url,
            headers=headers_json,
            json=payload,
            timeout=30
        )

        print(
            "META QR SEND:",
            send_response.status_code,
            send_response.text
        )

        if send_response.status_code in [200, 201]:
            guardar_mensaje(
                numero,
                "Terrace",
                "🧾 Código QR para transferencia enviado",
                "out"
            )
            return True

        print("ERROR ENVIANDO QR POR MEDIA_ID")
        return False

    except requests.RequestException as error:
        print("ERROR DE CONEXIÓN ENVIANDO QR:", str(error))
        return False

    except Exception as error:
        print("ERROR GENERAL ENVIANDO QR:", str(error))
        return False

@app.get("/whatsapp")
async def verificar_webhook(request: Request):
    params = request.query_params

    mode = params.get("hub.mode")
    token = params.get("hub.verify_token")
    challenge = params.get("hub.challenge")

    if mode == "subscribe" and token == META_VERIFY_TOKEN:
        return Response(content=challenge, media_type="text/plain")

    return Response(content="Token inválido", status_code=403)

@app.post("/whatsapp")
async def whatsapp(request: Request):

    body = await request.json()
    print("<<<<<<<<<<<< LLEGÓ UN WEBHOOK >>>>>>>>>>>>")
    print(json.dumps(body, indent=2, ensure_ascii=False))

    # Meta también manda eventos que no son mensajes, por ejemplo estados de entrega.
    # Si no hay mensaje nuevo, respondemos 200 para que Meta no reintente.
    try:
        value = body["entry"][0]["changes"][0]["value"]
        mensajes = value.get("messages", [])

        if not mensajes:
            return Response(content="EVENT_RECEIVED", media_type="text/plain")

        mensaje_obj = mensajes[0]

        message_id = mensaje_obj.get("id")
        numero = mensaje_obj.get("from")
        tipo_mensaje = mensaje_obj.get("type")

        if message_id and mensaje_fue_procesado(message_id):
            print("MENSAJE DUPLICADO IGNORADO:", message_id)

            return Response(
                content="EVENT_RECEIVED",
                media_type="text/plain"
            )

        if tipo_mensaje == "audio":
            nombre = value.get(
                "contacts",
                [{}]
            )[0].get("profile", {}).get("name", "")

            guardar_mensaje(
                numero,
                nombre,
                "🎤 El cliente envió un audio",
                "in"
            )

            enviar_texto(
                numero,
                "Solo recibimos mensajes de texto y llamadas al número de teléfono, "
                "no audios por WhatsApp."
            )

            return finalizar_webhook(message_id, numero)

        if tipo_mensaje != "text":
            enviar_texto(
                numero,
                "Solo recibimos mensajes de texto y llamadas al número de teléfono."
            )

            return finalizar_webhook(message_id, numero)

        mensaje = mensaje_obj["text"]["body"]
        nombre = value.get("contacts", [{}])[0].get("profile", {}).get("name", "")
        guardar_mensaje(numero, nombre, mensaje, "in")

    except Exception as e:
        print("ERROR LEYENDO WEBHOOK META:", e)
        return Response(
            content="EVENT_RECEIVED",
            media_type="text/plain"
        )

    mensaje_lower = mensaje.strip().lower()

    mensaje_normalizado = normalizar(mensaje)

    consultas_estado_pedido = [
        "se demora",
        "se demoran",
        "demora",
        "demoran",
        "demora mucho",
        "demoran mucho",
        "cuanto demora",
        "cuanto se demora",
        "cuanto demoran",
        "cuanto se demoran",
        "cuanto tarda",
        "cuanto tardan",
        "tarda",
        "tardan",
        "me confirmas cuanto tarda",
        "me confirma cuanto tarda",
        "ya vienen",
        "ya viene",
        "ya va",
        "ya salio",
        "ya sale",
        "como va el pedido",
        "como va mi pedido",
        "estado del pedido",
        "cuando llega",
        "cuando llegan",
        "cuando viene",
        "cuando vienen",
        "falta mucho",
        "cuanto falta",
        "en cuanto llega",
        "en cuanto tiempo llega"
    ]

    consulta_estado = any(
        frase in mensaje_normalizado
        for frase in consultas_estado_pedido
    )

    if consulta_estado:
        activar_atencion_humana(
            numero,
            minutos=30,
            motivo="Consulta sobre el estado o demora del pedido"
        )

        enviar_texto(
            numero,
            "Un momento por favor 😊 Un miembro del equipo revisará "
            "el estado de tu pedido y te responderá."
        )

        return finalizar_webhook(message_id, numero)

    mensajes_atencion_humana = {
        "humano",
        "persona",
        "asesor",
        "asesora",
        "empleado",
        "empleada",
        "agente",
        "hablar con humano",
        "hablar con un humano",
        "hablar con persona",
        "hablar con una persona",
        "quiero hablar con alguien",
        "quiero hablar con un asesor",
        "quiero hablar con una persona",
        "necesito ayuda humana",
        "atencion humana",
        "servicio al cliente"
    }

    solicita_humano = (
        mensaje_normalizado in mensajes_atencion_humana
        or "hablar con un humano" in mensaje_normalizado
        or "hablar con una persona" in mensaje_normalizado
        or "hablar con alguien" in mensaje_normalizado
        or "quiero un asesor" in mensaje_normalizado
    )

    if solicita_humano:
        activar_atencion_humana(
            numero,
            minutos=30,
            motivo="El cliente solicitó atención humana"
        )

        enviar_texto(
            numero,
            "Claro 😊 He solicitado la ayuda de una persona del equipo.\n\n"
            "El asistente automático quedará pausado durante los próximos "
            "30 minutos mientras un miembro de Terrace revisa tu conversación."
        )

        return finalizar_webhook(message_id, numero)
    
    if atencion_humana_activa(numero):
        return finalizar_webhook(message_id, numero)

    mensajes_cancelacion = {
        "cancelar",
        "cancelar pedido",
        "cancela el pedido",
        "cancela mi pedido",
        "quiero cancelar",
        "quiero cancelar el pedido",
        "ya no quiero el pedido",
        "ya no voy a pedir",
        "no voy a pedir",
        "olvida el pedido",
        "anula el pedido",
        "anular pedido"
    }

    pedido_actual = obtener_conversacion(numero)

    if mensaje_lower in mensajes_cancelacion:
        pedido_en_proceso = (
            bool(pedido_actual.get("items"))
            or bool(pedido_actual.get("nombre"))
            or bool(pedido_actual.get("tipo_entrega"))
            or bool(pedido_actual.get("direccion"))
            or bool(pedido_actual.get("metodo_pago"))
            or pedido_actual.get("esperando_confirmacion", False)
            or pedido_actual.get("esperando_metodo_pago", False)
        )

        borrar_conversacion(numero)

        if pedido_en_proceso:
            enviar_texto(
                numero,
                "Tu pedido fue cancelado correctamente ❌\n\n"
                "Cuando quieras hacer uno nuevo, solo escríbeme."
            )
        else:
            enviar_texto(
                numero,
                "No tienes ningún pedido en proceso.\n\n"
                "Cuando quieras hacer uno, solo escríbeme 😊"
            )

        return finalizar_webhook(message_id, numero)

    respuestas_confirmacion = {
        "si", "sí", "confirmo", "confirmado", "correcto",
        "esta bien", "está bien", "todo bien", "de acuerdo",
        "ok", "okay", "listo", "sii", "siii", "confirmar", "confirmalo"
    }

    respuestas_cambio = {
        "no", "cambiar", "cambio", "quiero cambiar",
        "hacer un cambio", "modificar", "corregir"
    }

    # Si ya se mostró el resumen, esta respuesta decide si se confirma
    # o si el cliente quiere modificar el pedido.
    if pedido_actual.get("esperando_confirmacion"):
        if mensaje_lower in respuestas_confirmacion:
            pedido_actual["esperando_confirmacion"] = False
            pedido_actual["pedido_confirmado"] = True
            pedido_actual["esperando_metodo_pago"] = False
            pedido_actual["metodo_pago"] = ""
            pedido_actual["fecha_confirmacion"] = hora_local()

            guardar_pedido_db(numero, pedido_actual)

            enviar_texto(
                numero,
                "✅ Listo. Tu pedido fue recibido."
            )

            borrar_conversacion(numero)

            return finalizar_webhook(message_id, numero)

        if mensaje_lower in respuestas_cambio:
            pedido_actual["esperando_confirmacion"] = False
            pedido_actual["pedido_confirmado"] = False
            guardar_conversacion(numero, pedido_actual)

            enviar_texto(
                numero,
                "Claro 😊 ¿Qué deseas cambiar del pedido?"
            )
            return finalizar_webhook(message_id, numero)

        # Si escribe directamente el cambio, dejamos que OpenAI lo procese.
        pedido_actual["esperando_confirmacion"] = False
        pedido_actual["pedido_confirmado"] = False
        guardar_conversacion(numero, pedido_actual)

    if mensaje_lower == "1":
        enviar_texto(numero, "Claro, aquí tienes nuestro menú 📋")
        enviar_menu(numero)
        return finalizar_webhook(message_id, numero)

    if mensaje_lower == "2":
        enviar_texto(
            numero,
            "Perfecto 😊 ¿Qué productos deseas ordenar y en qué cantidades?\n\n"
            "Puedes cancelar el pedido antes de confirmarlo escribiendo "
            "*cancelar* o *cancelar pedido*.\n\n"
            "Si necesitas atención personal, escribe *humano*."
        )
        return finalizar_webhook(message_id, numero)

    if "menu" in mensaje_lower or "menú" in mensaje_lower:
        enviar_texto(numero, "Claro, aquí tienes nuestro menú 📋")
        enviar_menu(numero)
        return finalizar_webhook(message_id, numero)

    pedido_actual = obtener_conversacion(numero)

    # contacto = obtener_contacto(numero)

    # if contacto and contacto.get("nombre"):
    #     pedido_actual["nombre"] = contacto["nombre"]

    try:
        inicio_openai = time.time()

        print(
            f"INICIANDO OPENAI | teléfono={numero} | "
            f"mensaje={mensaje}"
        )

        menu_activo = obtener_menu(solo_activos=True)

        completion = client.chat.completions.create(
            model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": f"""
    Eres el agente virtual de pedidos de Terrace.

    Debes actualizar el pedido usando:
    1. El pedido actual.
    2. El nuevo mensaje del cliente.
    3. El menú disponible.

    Devuelve siempre el pedido completo en JSON válido.

    MENÚ DISPONIBLE:

    {json.dumps(menu_activo, ensure_ascii=False)}

    ESTRUCTURA DE SALIDA:

    {{
    "tipo": "",
    "respuesta": "",
    "nombre": "",
    "items": [],
    "productos_ambiguos": [],
    "productos_no_disponibles": [],
    "tipo_entrega": "",
    "direccion": "",
    "hora": "",
    "metodo_pago": "",
    "notas": ""
    }}

    "tipo" solo puede ser:

    - "saludo"
    - "pregunta"
    - "menu"
    - "pedido"

    REGLAS DE INTENCIÓN:

    - Si el mensaje es solamente un saludo, usa tipo="saludo".
    - Si pide el menú, usa tipo="menu".
    - Si hace una pregunta informativa, usa tipo="pregunta".
    - Si agrega o corrige cualquier información del pedido, usa tipo="pedido".

    También es tipo="pedido" cuando el cliente responde solamente con:

    - su nombre;
    - el tipo de entrega;
    - la dirección o local;
    - una nota;
    - una cantidad;
    - una corrección.

    Ejemplos:

    "Hola" → saludo
    "Hola, quiero un latte" → pedido
    "¿A qué hora cierran?" → pregunta
    "Menú por favor" → menu
    "Zona Múltiple" → pedido
    "L3-26" → pedido

    REGLAS GENERALES:

    - Conserva toda la información válida del pedido actual.
    - No reemplaces datos existentes por valores vacíos.
    - Si el cliente corrige algo, actualiza solamente ese campo.
    - Si agrega un producto, conserva los productos anteriores.
    - Si elimina o cambia un producto explícitamente, actualízalo.
    - No inventes información.
    - Devuelve únicamente JSON válido.
    - Todos los campos deben estar presentes.

    PRODUCTOS:

    "items" debe tener esta estructura:

    [
    {{
        "producto": "Limonada de coco",
        "cantidad": "2",
        "notas": ""
    }}
    ]

    - Solo acepta productos que estén en el menú.
    - Usa el nombre oficial del producto del menú.
    - Si no indica cantidad, usa "1".
    - Reconoce cantidades escritas con números o palabras.
    - Si pide "tinto", usa el producto exacto "Tinto".
    - Las instrucciones específicas de un producto van en sus notas.

    Ejemplo:

    "Un latte con leche deslactosada y sin azúcar"

    {{
    "producto": "Latte",
    "cantidad": "1",
    "notas": "leche deslactosada, sin azúcar"
    }}

    PRODUCTOS AMBIGUOS:

    Si pide una categoría general que realmente exista en el menú y tenga varias opciones,
    no elijas una arbitrariamente.

    Solo puede considerarse ambiguo un término si existen dos o más productos del menú
    cuyos nombres contienen claramente ese término.

    Por ejemplo:
    - "limonada" es ambiguo si existen varias limonadas en el menú.
    - "malteada" es ambiguo si existen varias malteadas en el menú.
    - "capuccino" es ambiguo si existen varias opciones de capuccino.

    Ejemplos:

    - limonada
    - jugo
    - malteada
    - capuccino
    - sandwich
    - torta

    Agrégala a "productos_ambiguos" y no a "items".

    Si luego aclara la opción, elimina la ambigüedad y agrega el producto exacto.

    IMPORTANTE:

    "postre" no existe como categoría ni como producto en el menú.
    Si el cliente pide "postre", "postres", "postre de leche" o cualquier postre
    que no aparezca exactamente en el menú, agrégalo a
    "productos_no_disponibles", no a "productos_ambiguos".

    PRODUCTOS NO DISPONIBLES:

    Si pide un producto que no existe en el menú:

    - no lo agregues a "items";
    - agrégalo a "productos_no_disponibles";
    - no lo reemplaces por otro producto.

    TIPO DE ENTREGA:

    TIPO DE ENTREGA:

    "tipo_entrega" solo puede ser:

    - "recoger"
    - "domicilio"
    - ""

    IMPORTANTE:

    Si el cliente responde únicamente con un número o nombre de local,
    debe asumirse que el pedido es para llevar al local.

    Ejemplos:

    "Local 26"
    → tipo_entrega = "domicilio"
    → direccion = "Local 26"

    "26"
    → tipo_entrega = "domicilio"
    → direccion = "Local 26"

    "L3-26"
    → tipo_entrega = "domicilio"
    → direccion = "L3-26"

    "Zona Múltiple"
    → tipo_entrega = "domicilio"
    → direccion = "Zona Múltiple"

    "Tecni Play Sur"
    → tipo_entrega = "domicilio"
    → direccion = "Tecni Play Sur"

    Siempre que el mensaje sea únicamente una ubicación o un local,
    interpreta que el cliente respondió a la pregunta
    "¿Es para llevar a un local o pasas a recogerlo?"

    Si el mensaje del cliente contiene únicamente un local o una ubicación,
    además de llenar "direccion", debes establecer automáticamente:

    "tipo_entrega": "domicilio"

    aunque el cliente nunca haya escrito la palabra "domicilio".

    Usa "domicilio" cuando el cliente escriba expresiones como:

    - llevar
    - para llevar
    - es para llevar
    - me lo llevan
    - que me lo lleven
    - domicilio
    - a domicilio
    - enviar al local
    - llevar al local

    Si el cliente responde solamente "llevar" o "para llevar":

    - establece "tipo_entrega": "domicilio";
    - conserva "direccion": "" si todavía no indicó el local;
    - no inventes ningún número o nombre de local.

    La aplicación debe preguntarle después cuál es el número o nombre del local.

    Usa "recoger" para expresiones como:

    - recoger
    - lo recojo
    - paso por él
    - pickup

    Si es "recoger", la dirección puede quedar vacía.

    Si es "domicilio", la dirección es obligatoria.

    DIRECCIÓN O LOCAL:

    En este negocio, "direccion" es el lugar dentro del centro comercial donde se entrega el pedido.

    Puede ser:

    - un código de local;
    - el nombre de un establecimiento;
    - ambos;
    - una referencia clara.

    Ejemplos válidos:

    - "L3-26"
    - "P4-24(P17)"
    - "Zona Múltiple"
    - "Tecni Play Sur"
    - "Mar y Juancho"
    - "Bodega"
    - "Zona Múltiple, local L1-103"

    No es obligatorio que el cliente diga "local" o "dirección".

    Ejemplos:

    "L3-26"
    → "direccion": "L3-26"

    "Estoy en Tecni Play Sur"
    → "direccion": "Tecni Play Sur"

    "Llévalo a Bodega"
    → "direccion": "Bodega"

    "Zona Múltiple, local L1-103"
    → "direccion": "Zona Múltiple, local L1-103"

    No confundas el nombre o código del local con:

    - el nombre del cliente;
    - un producto;
    - una cantidad;
    - una hora;
    - una nota.

    NOMBRE:

    Guarda como "nombre" solamente el nombre de la persona.

    Ejemplos:

    - "Soy Carlos"
    - "A nombre de Andrea"
    - "Ponlo a nombre de Juan"

    No guardes nombres de establecimientos en "nombre".

    Los campos "productos_ambiguos" y "productos_no_disponibles"
    describen únicamente el NUEVO MENSAJE DEL CLIENTE.

    Nunca conserves esos valores del pedido anterior.
    Si el nuevo mensaje no contiene productos ambiguos o no disponibles,
    devuelve esos campos como listas vacías.

    HORA Y NOTAS:

    - La hora no es obligatoria.
    - Las notas generales van en "notas".
    - Las instrucciones de un producto van en las notas del item.

    INFORMACIÓN DEL RESTAURANTE:

    Nombre: Terrace
    Horario: 8:00 AM a 10:00 PM
    Métodos de pago: efectivo y transferencia

    Si no sabes la respuesta a una pregunta, no inventes información.

    Antes de responder, verifica:

    1. Que el JSON sea válido.
    2. Que no hayas eliminado información válida.
    3. Que los productos existan en el menú.
    4. Que un nombre o código de local esté en "direccion".
    5. Que "tipo_entrega" tenga un valor permitido.

    IMPORTANTE SOBRE LA INTENCIÓN:

    La intención debe determinarse solamente usando el NUEVO MENSAJE DEL CLIENTE.

    No copies ni conserves el valor anterior de "tipo" o "respuesta".

    Los campos "tipo" y "respuesta" son temporales y describen solamente el mensaje actual.

    Si el nuevo mensaje contiene productos, cantidades, información de entrega,
    dirección, local, nombre, notas o correcciones,
    siempre devuelve tipo="pedido", aunque el pedido anterior haya sido un saludo.

    Ejemplo:

    Pedido anterior:
    El cliente solamente dijo "Hola".

    Nuevo mensaje:
    "dame 2 empanadas, 1 café y 3 capuchinos"

    Resultado:
    tipo="pedido"

    Nunca devuelvas tipo="saludo" cuando el nuevo mensaje contenga productos o cantidades.
    """
            },
                {
                    "role": "user",
                    "content": f"""
    Pedido actual:

    {json.dumps(pedido_actual, ensure_ascii=False)}

    Nuevo mensaje del cliente:

    {mensaje}

    Devuelve el pedido completo y actualizado usando exactamente esta estructura:

    {{
    "tipo": "",
    "respuesta": "",
    "nombre": "",
    "items": [],
    "productos_ambiguos": [],
    "productos_no_disponibles": [],
    "tipo_entrega": "",
    "direccion": "",
    "hora": "",
    "metodo_pago": "",
    "notas": ""
    }}

    Devuelve cualquier respuesta para el cliente dentro del campo "respuesta".

    No escribas texto fuera del JSON.

    Para tipo="saludo", tipo="pregunta" o tipo="menu", usa el campo "respuesta" cuando corresponda.

    Para tipo="pedido", puedes usar "respuesta" si necesitas dar una confirmación breve, pero la aplicación también puede decidir qué pregunta hacer después.
    """
                }
            ],
            response_format={"type": "json_object"},
            max_tokens=1000,
            temperature=0
        )

        duracion_openai = time.time() - inicio_openai

        print(
            f"OPENAI TERMINÓ EN {duracion_openai:.2f} SEGUNDOS "
            f"| teléfono={numero}"
        )

        data = completion.choices[0].message.content
        print("IA:", data)

        pedido = json.loads(data)
        pedido = validar_items_pedido(pedido)

        respuestas_domicilio = {
            "llevar",
            "para llevar",
            "es para llevar",
            "domicilio",
            "a domicilio",
            "me lo llevan",
            "que me lo lleven"
        }

        if mensaje_normalizado in respuestas_domicilio:
            pedido["tipo_entrega"] = "domicilio"

            # Si todavía no indicó un local, se deja vacío
            # para que el sistema se lo pregunte.
            if not pedido.get("direccion"):
                pedido["direccion"] = ""

        pedido["pedido_confirmado"] = pedido_actual.get(
            "pedido_confirmado",
            False
        )

        pedido["fecha_confirmacion"] = pedido_actual.get(
            "fecha_confirmacion",
            ""
        )

        pedido["esperando_confirmacion"] = pedido_actual.get(
            "esperando_confirmacion",
            False
        )

        pedido["esperando_metodo_pago"] = pedido_actual.get(
            "esperando_metodo_pago",
            False
        )

    except Exception as e:
        print("ERROR OPENAI:", repr(e))
        traceback.print_exc()

        enviar_texto(
            numero,
            "Lo siento, ocurrió un error procesando tu pedido. "
            "¿Podrías intentarlo nuevamente?"
        )

        return Response(
            content="EVENT_RECEIVED",
            media_type="text/plain"
        )

    # contacto = obtener_contacto(numero)

    # if contacto and contacto.get("nombre") and not pedido.get("nombre"):
    #     pedido["nombre"] = contacto["nombre"]

    guardar_conversacion(numero, pedido)

    tipo = pedido.get("tipo", "").lower()

    if tipo == "saludo":
        enviar_texto(
            numero,
            "¡Hola! Bienvenido a Terrace ☕🍰\n\n"
            "Por favor, escribe el número de tu opción:\n\n"
            "1. Ver el menú\n"
            "2. Hacer un pedido\n\n"
            "También puedes escribir *humano* en cualquier momento "
            "si deseas hablar con una persona."
        )

        return finalizar_webhook(message_id, numero)

    if tipo == "menu":
        enviar_texto(numero, "Claro, aquí tienes nuestro menú 📋")
        enviar_menu(numero)
        return finalizar_webhook(message_id, numero)

    if tipo == "pregunta":
        enviar_texto(numero, pedido.get("respuesta", ""))
        return finalizar_webhook(message_id, numero)

    faltantes = []
    productos_ambiguos = pedido.get("productos_ambiguos", [])
    productos_no_disponibles = pedido.get("productos_no_disponibles", [])

    tipo_entrega = pedido.get("tipo_entrega", "").strip().lower()

    for campo in CAMPOS_OBLIGATORIOS:
        if campo == "direccion":
            if tipo_entrega == "recoger":
                continue

            if tipo_entrega == "domicilio" and not pedido.get("direccion"):
                faltantes.append("direccion")
                continue

        if campo == "items":
            if len(pedido.get("items", [])) == 0:
                faltantes.append("items")
            continue

        if not pedido.get(campo):
            faltantes.append(campo)

    if productos_no_disponibles:
        enviar_texto(
            numero,
            "Lo siento, no tenemos disponible: "
            + ", ".join(productos_no_disponibles)
            + ". Puedes elegir otro producto del menú."
        )

        # El error ya fue informado, por lo tanto no debe quedarse
        # guardado para los siguientes mensajes.
        pedido["productos_no_disponibles"] = []
        guardar_conversacion(numero, pedido)

        return finalizar_webhook(message_id, numero)

    if productos_ambiguos:
        mensajes = []

        nombres_ambiguos_normalizados = {
            normalizar(producto)
            for producto in productos_ambiguos
        }

        pedido["items"] = [
            item
            for item in pedido.get("items", [])
            if normalizar(item.get("producto", ""))
            not in nombres_ambiguos_normalizados
        ]

        for ambiguo in productos_ambiguos:
            opciones = opciones_por_producto(ambiguo)

            if opciones:
                mensajes.append(
                    f"¿Cuál opción deseas de {ambiguo}? Tenemos: "
                    + ", ".join(opciones)
                )
            else:
                mensajes.append(
                    f"¿Cuál opción deseas de {ambiguo}?"
                )

        enviar_texto(numero, "\n\n".join(mensajes))

        # La ambigüedad ya fue informada. No debe bloquear
        # permanentemente los siguientes mensajes.
        pedido["productos_ambiguos"] = []
        guardar_conversacion(numero, pedido)

        return finalizar_webhook(message_id, numero)

    if len(faltantes) == 0:
        items_resumen = []

        for item in pedido.get("items", []):
            cantidad = item.get("cantidad", "1")
            producto = item.get("producto", "")
            notas_item = item.get("notas", "").strip()

            texto_item = f"{cantidad} x {producto}"

            if notas_item:
                texto_item += f" ({notas_item})"

            items_resumen.append(texto_item)

        items_texto = "\n".join(items_resumen)

        total = calcular_total(pedido)

        # 1. Primero se muestra el resumen y se pide confirmación.
        if not pedido.get("pedido_confirmado", False):
            pedido["esperando_confirmacion"] = True
            pedido["esperando_metodo_pago"] = False
            guardar_conversacion(numero, pedido)

            entrega_texto = (
                "Recoger en Terrace"
                if pedido.get("tipo_entrega") == "recoger"
                else pedido.get("direccion", "")
            )

            enviar_texto(
                numero,
                "Por favor confirma tu pedido:\n\n"
                f"👤 Nombre: {pedido.get('nombre', '')}\n"
                f"🛒 Pedido: {items_texto}\n"
                f"📍 Entrega: {entrega_texto}\n"
                f"💰 Total: ${total:,} COP\n\n"
                "¿Está correcto? Responde *sí* para confirmar "
                "o indícame qué deseas cambiar.\n\n"
                "También puedes escribir *cancelar* o *cancelar pedido* "
                "para detener el proceso."
            )

            return finalizar_webhook(message_id, numero)
        

    else:
        campo = faltantes[0]

        preguntas = {
            "nombre": "Perfecto. ¿A nombre de quien la orden?",
            "items": "¿Qué productos deseas ordenar y en qué cantidades?",
            "tipo_entrega": "¿Es para llevar a un local o pasas a recogerlo?",
            "direccion": "¿Cuál es el número o nombre del local?"
        }

        enviar_texto(numero, preguntas.get(campo, "Me falta información para completar tu pedido."))

    return finalizar_webhook(message_id, numero)

@app.post("/manual/send")
async def enviar_manual(request: Request):
    form = await request.form()

    telefono = str(form.get("telefono", "")).strip()
    mensaje = str(form.get("mensaje", "")).strip()

    if not telefono or not mensaje:
        return JSONResponse(
            {
                "success": False,
                "error": "Faltan el teléfono o el mensaje."
            },
            status_code=400
        )

    enviado = enviar_texto(telefono, mensaje)

    if not enviado:
        return JSONResponse(
            {
                "success": False,
                "error": "Meta no pudo enviar el mensaje."
            },
            status_code=502
        )

    activar_atencion_humana(
        telefono,
        minutos=30,
        motivo="Respuesta manual del personal"
    )

    return {
        "success": True,
        "human_mode": True,
        "minutes": 30
    }

@app.get("/orders/manual")
async def mostrar_formulario_pedido_manual(request: Request):
    menu_activo = obtener_menu(solo_activos=True)

    return templates.TemplateResponse(
        request=request,
        name="manual_order.html",
        context={
            "menu": menu_activo,
            "error": None
        }
    )


@app.post("/orders/manual")
async def crear_pedido_manual(request: Request):
    form = await request.form()

    nombre = str(form.get("nombre", "")).strip()
    telefono = str(form.get("telefono", "")).strip()
    tipo_entrega = str(form.get("tipo_entrega", "")).strip().lower()
    direccion = str(form.get("direccion", "")).strip()
    hora = str(form.get("hora", "")).strip()
    metodo_pago = ""
    notas = str(form.get("notas", "")).strip()

    productos = form.getlist("producto")
    cantidades = form.getlist("cantidad")
    notas_productos = form.getlist("notas_producto")

    errores = []

    if not nombre:
        errores.append("Debes ingresar el nombre del cliente.")

    if tipo_entrega not in {"domicilio", "recoger"}:
        errores.append("Debes seleccionar el tipo de entrega.")

    if tipo_entrega == "domicilio" and not direccion:
        errores.append(
            "Debes ingresar el local o dirección para el domicilio."
        )

    items = []

    for indice, nombre_producto in enumerate(productos):
        nombre_producto = str(nombre_producto).strip()

        if not nombre_producto:
            continue

        producto_menu = buscar_producto(nombre_producto)

        if not producto_menu:
            errores.append(
                f"El producto '{nombre_producto}' no se encuentra en el menú."
            )
            continue

        try:
            cantidad = int(cantidades[indice])
        except (ValueError, TypeError, IndexError):
            cantidad = 1

        if cantidad < 1:
            errores.append(
                f"La cantidad de {producto_menu['producto']} debe ser mayor que cero."
            )
            continue

        nota_producto = ""

        if indice < len(notas_productos):
            nota_producto = str(notas_productos[indice]).strip()

        items.append({
            "producto": producto_menu["producto"],
            "cantidad": str(cantidad),
            "notas": nota_producto
        })

    if not items:
        errores.append("Debes agregar al menos un producto.")

    if errores:
        return templates.TemplateResponse(
            request=request,
            name="manual_order.html",
            context={
                "menu": obtener_menu(solo_activos=True),
                "error": " ".join(errores),
                "form_data": {
                    "nombre": nombre,
                    "telefono": telefono,
                    "tipo_entrega": tipo_entrega,
                    "direccion": direccion,
                    "hora": hora,
                    "notas": notas
                }
            },
            status_code=400
        )

    pedido = {
        "nombre": nombre,
        "items": items,
        "tipo_entrega": tipo_entrega,
        "direccion": direccion if tipo_entrega == "domicilio" else "",
        "hora": hora,
        "metodo_pago": "",
        "notas": notas
    }

    numero_pedido = telefono if telefono else "PEDIDO MANUAL"

    guardar_pedido_db(numero_pedido, pedido)

    return RedirectResponse(
        url="/dashboard",
        status_code=303
    )

@app.get("/menu/admin")
async def administrar_menu(request: Request):
    productos = obtener_menu(solo_activos=False)

    return templates.TemplateResponse(
        request=request,
        name="admin_menu.html",
        context={
            "productos": productos,
            "error": None
        }
    )


@app.post("/menu/admin/add")
async def agregar_producto(request: Request):
    form = await request.form()

    producto = str(form.get("producto", "")).strip()
    precio_texto = str(form.get("precio", "")).strip()

    if not producto:
        raise HTTPException(
            status_code=400,
            detail="El nombre del producto es obligatorio."
        )

    try:
        precio = int(precio_texto)
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail="El precio debe ser un número entero."
        )

    if precio < 0:
        raise HTTPException(
            status_code=400,
            detail="El precio no puede ser negativo."
        )

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            INSERT INTO products (
                producto,
                precio,
                activo
            )
            VALUES (%s, %s, TRUE)
        """, (
            producto,
            precio
        ))

        conn.commit()

    except psycopg2.IntegrityError:
        conn.rollback()
        conn.close()

        raise HTTPException(
            status_code=409,
            detail="Ya existe un producto con ese nombre."
        )

    conn.close()

    return RedirectResponse(
        url="/menu/admin",
        status_code=303
    )


@app.post("/menu/admin/{product_id}/edit")
async def editar_producto(product_id: int, request: Request):
    form = await request.form()

    producto = str(form.get("producto", "")).strip()
    precio_texto = str(form.get("precio", "")).strip()

    if not producto:
        raise HTTPException(
            status_code=400,
            detail="El nombre del producto es obligatorio."
        )

    try:
        precio = int(precio_texto)
    except ValueError:
        raise HTTPException(
            status_code=400,
            detail="El precio debe ser un número entero."
        )

    if precio < 0:
        raise HTTPException(
            status_code=400,
            detail="El precio no puede ser negativo."
        )

    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            UPDATE products
            SET
                producto = %s,
                precio = %s,
                actualizado_en = CURRENT_TIMESTAMP
            WHERE id = %s
        """, (
            producto,
            precio,
            product_id
        ))

        conn.commit()

    except psycopg2.IntegrityError:
        conn.rollback()
        conn.close()

        raise HTTPException(
            status_code=409,
            detail="Ya existe otro producto con ese nombre."
        )

    conn.close()

    return RedirectResponse(
        url="/menu/admin",
        status_code=303
    )


@app.post("/menu/admin/{product_id}/toggle")
async def cambiar_disponibilidad_producto(product_id: int):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE products
        SET
            activo = NOT activo,
            actualizado_en = CURRENT_TIMESTAMP
        WHERE id = %s
    """, (product_id,))

    conn.commit()
    conn.close()

    return RedirectResponse(
        url="/menu/admin",
        status_code=303
    )

@app.get("/dashboard")
async def dashboard(request: Request):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT * FROM orders
        WHERE estado != 'entregado'
        ORDER BY id DESC
    """)

    orders = cursor.fetchall()
    orders_with_items = []

    for order in orders:
        cursor.execute("""
            SELECT producto, cantidad, notas
            FROM order_items
            WHERE order_id = %s
        """, (order["id"],))

        items = cursor.fetchall()

        orders_with_items.append({
            "order": order,
            "items": items
        })

    conn.close()

    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "orders": orders_with_items
        }
    )


@app.post("/orders/{order_id}/estado/{estado}")
async def cambiar_estado(order_id: int, estado: str):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        "UPDATE orders SET estado = %s WHERE id = %s",
        (estado, order_id)
    )

    conn.commit()
    conn.close()

    return RedirectResponse(url="/dashboard", status_code=303)

@app.post("/orders/{order_id}/cancel")
async def cancelar_pedido(order_id: int):

    conn = get_connection()
    cursor = conn.cursor()

    # Primero borrar los productos
    cursor.execute("""
        DELETE FROM order_items
        WHERE order_id = %s
    """, (order_id,))

    # Después borrar el pedido
    cursor.execute("""
        DELETE FROM orders
        WHERE id = %s
    """, (order_id,))

    conn.commit()
    conn.close()

    return RedirectResponse(
        url="/dashboard",
        status_code=303
    )

@app.get("/export/excel")
async def export_excel(fecha: str = None):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    if fecha:
        cursor.execute("""
            SELECT 
                o.id AS order_id,
                o.fecha,
                o.telefono,
                o.nombre,
                oi.producto,
                oi.cantidad,
                oi.notas AS notas_producto,
                o.tipo_entrega,
                o.direccion,
                o.hora,
                o.metodo_pago,
                o.notas AS notas_pedido,
                o.estado,
                o.total
            FROM orders o
            LEFT JOIN order_items oi ON o.id = oi.order_id
            WHERE o.fecha LIKE %s
            ORDER BY o.id DESC
        """, (f"{fecha}%",))
    else:
        cursor.execute("""
            SELECT 
                o.id AS order_id,
                o.fecha,
                o.telefono,
                o.nombre,
                oi.producto,
                oi.cantidad,
                oi.notas AS notas_producto,
                o.tipo_entrega,
                o.direccion,
                o.hora,
                o.metodo_pago,
                o.notas AS notas_pedido,
                o.estado,
                o.total
            FROM orders o
            LEFT JOIN order_items oi ON o.id = oi.order_id
            ORDER BY o.id DESC
        """)

    rows = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos"

    headers = [
        "Order ID",
        "Fecha",
        "Teléfono",
        "Nombre",
        "Producto",
        "Cantidad",
        "Notas producto",
        "Tipo entrega",
        "Dirección",
        "Hora",
        "Método pago",
        "Notas pedido",
        "Estado",
        "Total"
    ]

    ws.append(headers)

    for row in rows:
        ws.append([
            row["order_id"],
            row["fecha"],
            row["telefono"],
            row["nombre"],
            row["producto"],
            row["cantidad"],
            row["notas_producto"],
            row["tipo_entrega"],
            row["direccion"],
            row["hora"],
            row["metodo_pago"],
            row["notas_pedido"],
            row["estado"],
            row["total"]
        ])

    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=pedidos_terrace.xlsx"
        }
    )

@app.post("/clear/messages")
async def clear_messages():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM messages")
    cursor.execute("DELETE FROM conversation_states")
    cursor.execute("DELETE FROM human_handoffs")

    conn.commit()
    conn.close()

    return {
        "status": "Chats, conversaciones y controles humanos borrados"
    }

@app.post("/chat/{telefono}/human/activate")
async def activar_humano_desde_chat(telefono: str):
    activar_atencion_humana(
        telefono,
        minutos=30,
        motivo="Activado manualmente desde el panel"
    )

    return {
        "success": True,
        "activo": True,
        "minutos": 30
    }


@app.post("/chat/{telefono}/human/deactivate")
async def desactivar_humano_desde_chat(telefono: str):
    desactivar_atencion_humana(telefono)

    return {
        "success": True,
        "activo": False
    }


@app.get("/chat/{telefono}/human/status")
async def estado_humano_chat(telefono: str):
    estado = obtener_estado_atencion_humana(telefono)

    activo_hasta = estado.get("activo_hasta")

    if activo_hasta:
        activo_hasta = activo_hasta.isoformat()

    return {
        "activo": bool(estado.get("activo")),
        "activo_hasta": activo_hasta,
        "motivo": estado.get("motivo", "")
    }

@app.get("/privacy", response_class=HTMLResponse)
async def privacy_policy():
    return """
    <html>
    <head><title>Privacy Policy - Terrace Coffee</title></head>
    <body style="font-family: Arial; max-width: 800px; margin: 40px auto; line-height: 1.6;">
        <h1>Privacy Policy</h1>
        <p>Last updated: July 2026</p>

        <p>Terrace Coffee uses WhatsApp to communicate with customers, receive orders, answer questions, and provide customer support.</p>

        <h2>Information We Collect</h2>
        <p>We may collect your name, phone number, order details, delivery information, payment preference, and any message you send to us through WhatsApp.</p>

        <h2>How We Use Information</h2>
        <p>We use this information only to process orders, respond to customer requests, provide service updates, and improve our customer experience.</p>

        <h2>Data Sharing</h2>
        <p>We do not sell or rent customer personal information. Information may only be shared with service providers when necessary to operate our ordering and messaging system.</p>

        <h2>Data Retention</h2>
        <p>We keep customer information only as long as necessary for order management, customer support, legal, or business purposes.</p>

        <h2>Contact</h2>
        <p>If you have questions about this Privacy Policy, contact us at: juandavidvasquezescobar2020@gmail.com</p>
    </body>
    </html>
    """


@app.get("/terms", response_class=HTMLResponse)
async def terms_of_service():
    return """
    <html>
    <head><title>Terms of Service - Terrace Coffee</title></head>
    <body style="font-family: Arial; max-width: 800px; margin: 40px auto; line-height: 1.6;">
        <h1>Terms of Service</h1>
        <p>Last updated: July 2026</p>

        <p>By using Terrace Coffee's WhatsApp ordering service, you agree to these Terms of Service.</p>

        <h2>Service</h2>
        <p>Our WhatsApp service allows customers to ask questions, view menu information, place orders, and communicate with Terrace Coffee.</p>

        <h2>Orders</h2>
        <p>Order availability, prices, delivery options, and preparation times may vary. Terrace Coffee may confirm, modify, or cancel orders when necessary.</p>

        <h2>User Responsibilities</h2>
        <p>Customers agree to provide accurate information when placing an order, including name, phone number, delivery details, and payment preference.</p>

        <h2>Limitation of Liability</h2>
        <p>Terrace Coffee is not responsible for delays, incorrect information submitted by customers, or service interruptions outside our control.</p>

        <h2>Contact</h2>
        <p>For questions about these Terms, contact us at: juandavidvasquezescobar2020@gmail.com</p>
    </body>
    </html>
    """


@app.get("/data-deletion", response_class=HTMLResponse)
async def data_deletion():
    return """
    <html>
    <head><title>Data Deletion - Terrace Coffee</title></head>
    <body style="font-family: Arial; max-width: 800px; margin: 40px auto; line-height: 1.6;">
        <h1>User Data Deletion Instructions</h1>
        <p>Last updated: July 2026</p>

        <p>If you want Terrace Coffee to delete your personal data collected through our WhatsApp ordering service, please contact us by email.</p>

        <h2>How to Request Deletion</h2>
        <p>Send an email to: juandavidvasquezescobar2020@gmail.com</p>

        <p>Please include your WhatsApp phone number and write: "Delete my data" in the email subject.</p>

        <h2>Processing Time</h2>
        <p>We will review and process deletion requests within a reasonable time, unless we are required to keep certain information for legal, security, or business record purposes.</p>

        <h2>Contact</h2>
        <p>Email: juandavidvasquezescobar2020@gmail.com</p>
    </body>
    </html>
    """
@app.get("/inbox")
async def inbox(request: Request):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT 
            m.telefono,
            MAX(m.nombre) AS nombre,
            MAX(m.fecha::timestamp) AS ultima_fecha,
            COUNT(*) FILTER (
                WHERE m.direccion = 'in'
                AND m.leido = FALSE
            ) AS no_leidos,
            COALESCE(
                BOOL_OR(
                    h.activo_hasta IS NOT NULL
                    AND h.activo_hasta > NOW()
                ),
                FALSE
            ) AS atencion_humana
        FROM messages m
        LEFT JOIN human_handoffs h
            ON h.telefono = m.telefono
        GROUP BY m.telefono
        ORDER BY ultima_fecha DESC
    """)

    chats = cursor.fetchall()
    conn.close()

    return templates.TemplateResponse(
        request=request,
        name="inbox.html",
        context={
            "chats": chats
        }
    )

@app.get("/inbox/chats")
async def inbox_chats():
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT
            m.telefono,
            MAX(m.nombre) AS nombre,
            MAX(m.fecha::timestamp) AS ultima_fecha,
            COUNT(*) FILTER (
                WHERE m.direccion = 'in'
                AND m.leido = FALSE
            ) AS no_leidos,
            COALESCE(
                BOOL_OR(
                    h.activo_hasta IS NOT NULL
                    AND h.activo_hasta > NOW()
                ),
                FALSE
            ) AS atencion_humana
        FROM messages m
        LEFT JOIN human_handoffs h
            ON h.telefono = m.telefono
        GROUP BY m.telefono
        ORDER BY ultima_fecha DESC
    """)

    chats = cursor.fetchall()
    conn.close()

    html = ""

    for chat in chats:
        atencion_humana = bool(chat["atencion_humana"])

        chat_class = "chat chat-human" if atencion_humana else "chat"
        avatar_class = "avatar avatar-human" if atencion_humana else "avatar"

        badge = ""

        if chat["no_leidos"] > 0:
            badge_class = (
                "badge badge-human"
                if atencion_humana
                else "badge"
            )

            badge = f"""
            <div class="{badge_class}">
                {chat["no_leidos"]}
            </div>
            """

        human_label = ""

        if atencion_humana:
            human_label = """
            <div class="human-label">
                👤 Atención humana activa
            </div>
            """

        inicial = (chat["nombre"] or chat["telefono"])[0].upper()

        html += f"""
        <a class="{chat_class}" href="/chat/{chat['telefono']}">

            <div class="{avatar_class}">
                {inicial}
            </div>

            <div class="info">

                <div class="nombre">
                    {chat["nombre"] or chat["telefono"]}
                </div>

                <div class="telefono">
                    {chat["telefono"]}
                </div>

                <div class="fecha">
                    Último mensaje: {chat["ultima_fecha"]}
                </div>

                {human_label}

            </div>

            {badge}

        </a>
        """

    return HTMLResponse(content=html)

@app.get("/chat/{telefono}")
async def chat(request: Request, telefono: str):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        UPDATE messages
        SET leido = TRUE
        WHERE telefono = %s
          AND direccion = 'in'
          AND leido = FALSE
    """, (telefono,))

    conn.commit()

    cursor.execute("""
        SELECT *
        FROM messages
        WHERE telefono = %s
        ORDER BY id ASC
    """, (telefono,))

    mensajes = cursor.fetchall()
    conn.close()

    estado_humano = obtener_estado_atencion_humana(telefono)

    return templates.TemplateResponse(
        request=request,
        name="chat.html",
        context={
            "telefono": telefono,
            "mensajes": mensajes,
            "human_mode": bool(estado_humano.get("activo")),
            "human_until": estado_humano.get("activo_hasta"),
            "human_reason": estado_humano.get("motivo", "")
        }
    )

@app.get("/chat/{telefono}/messages")
async def chat_messages(telefono: str):
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT *
        FROM messages
        WHERE telefono = %s
        ORDER BY id ASC
    """, (telefono,))

    mensajes = cursor.fetchall()
    conn.close()

    html = ""

    for m in mensajes:
        if m["direccion"] == "out":
            style = "background:#d1fae5; margin-left:auto;"
        else:
            style = "background:#e5e7eb;"

        html += f"""
        <div style="
            margin:10px;
            padding:10px;
            border-radius:10px;
            max-width:60%;
            {style}
        ">
            <p>{m["mensaje"]}</p>
            <small>{m["fecha"]}</small>
        </div>
        """

    return HTMLResponse(content=html)