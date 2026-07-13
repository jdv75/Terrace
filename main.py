from fastapi import FastAPI, Request, HTTPException
import requests
from openai import OpenAI
from fastapi.responses import Response, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
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

with open("menu.json", "r", encoding="utf-8") as file:
    MENU = json.load(file)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
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
    "direccion",
    "metodo_pago"
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

    conn.commit()
    conn.close()

crear_tabla()

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
        "confirmar_direccion_guardada": ""
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
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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

    for item in pedido.get("items", []):
        producto = item.get("producto", "").lower()
        try:
            cantidad = int(item.get("cantidad", 1))
        except:
            cantidad = 1

        producto_encontrado = buscar_producto(producto)

        if producto_encontrado:
            precio = producto_encontrado["precio"]
            item["producto"] = producto_encontrado["producto"]
        else:
            precio = 0

        extra = 0
        notas_item = item.get("notas", "").lower()

        if "deslactosada" in notas_item:
            extra = 2000

        total += (precio + extra) * cantidad

    return total

def buscar_producto(nombre_producto):
    nombre_producto = normalizar(nombre_producto)

    # match exacto primero
    for item in MENU:
        if normalizar(item["producto"]) == nombre_producto:
            return item

    nombres_normalizados = [normalizar(item["producto"]) for item in MENU]

    resultado = process.extractOne(
        nombre_producto,
        nombres_normalizados,
        scorer=fuzz.WRatio
    )

    if resultado is None:
        return None

    nombre_encontrado, score, index = resultado

    if score >= 75:
        return MENU[index]

    return None

def opciones_por_producto(producto_ambiguo):
    producto_ambiguo = normalizar(producto_ambiguo)

    opciones = []
    for item in MENU:
        nombre = item["producto"]
        nombre_norm = normalizar(nombre)

        if producto_ambiguo in nombre_norm:
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
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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

    response = requests.post(url, headers=headers, json=payload)
    print("META RESPONSE:", response.status_code, response.text)
    if response.status_code in [200, 201]:
        guardar_mensaje(numero, "Terrace", texto, "out")


def enviar_documento(numero, link=None, filename="Menu_Terrace.pdf"):
    ruta_pdf = os.path.join("static", "menu.pdf")

    if not os.path.exists(ruta_pdf):
        print("ERROR: No existe el archivo:", ruta_pdf)
        return False

    headers_auth = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}"
    }

    # 1. Subir el PDF directamente a Meta
    upload_url = (
        f"https://graph.facebook.com/v25.0/"
        f"{META_PHONE_NUMBER_ID}/media"
    )

    try:
        with open(ruta_pdf, "rb") as archivo_pdf:
            files = {
                "file": (
                    filename,
                    archivo_pdf,
                    "application/pdf"
                )
            }

            data = {
                "messaging_product": "whatsapp",
                "type": "application/pdf"
            }

            upload_response = requests.post(
                upload_url,
                headers=headers_auth,
                files=files,
                data=data,
                timeout=60
            )

        print(
            "META MEDIA UPLOAD:",
            upload_response.status_code,
            upload_response.text
        )

        if upload_response.status_code not in [200, 201]:
            print("ERROR SUBIENDO PDF A META")
            return False

        media_id = upload_response.json().get("id")

        if not media_id:
            print("ERROR: Meta no devolvió media_id")
            return False

        # 2. Enviar el documento usando el media_id
        messages_url = (
            f"https://graph.facebook.com/v25.0/"
            f"{META_PHONE_NUMBER_ID}/messages"
        )

        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": numero,
            "type": "document",
            "document": {
                "id": media_id,
                "filename": filename,
                "caption": "Menú de Terrace 📋"
            }
        }

        headers_json = {
            "Authorization": f"Bearer {META_ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        send_response = requests.post(
            messages_url,
            headers=headers_json,
            json=payload,
            timeout=30
        )

        print(
            "META DOCUMENT SEND:",
            send_response.status_code,
            send_response.text
        )

        if send_response.status_code in [200, 201]:
            guardar_mensaje(
                numero,
                "Terrace",
                "📋 Menú de Terrace enviado",
                "out"
            )
            return True

        print("ERROR ENVIANDO PDF POR MEDIA_ID")
        return False

    except requests.RequestException as error:
        print("ERROR DE CONEXIÓN CON META:", str(error))
        return False

    except Exception as error:
        print("ERROR GENERAL ENVIANDO DOCUMENTO:", str(error))
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
        numero = mensaje_obj.get("from")
        tipo_mensaje = mensaje_obj.get("type")

        if tipo_mensaje != "text":
            enviar_texto(numero, "Por ahora solo puedo recibir mensajes de texto. Escríbeme tu pedido o pide el menú 😊")
            return Response(content="EVENT_RECEIVED", media_type="text/plain")

        mensaje = mensaje_obj["text"]["body"]
        nombre = value.get("contacts", [{}])[0].get("profile", {}).get("name", "")
        guardar_mensaje(numero, nombre, mensaje, "in")

    except Exception as e:
        print("ERROR LEYENDO WEBHOOK META:", e)
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    mensaje_lower = mensaje.strip().lower()

    pedido_actual = obtener_conversacion(numero)

    mensajes_transferencia = [
        "transferencia",
        "transferencia por favor",
        "pago por transferencia",
        "quiero pagar por transferencia",
        "qr",
        "codigo qr",
        "código qr",
        "mandame el qr",
        "mándame el qr"
    ]

    pedido_vacio = (
        not pedido_actual.get("items")
        and not pedido_actual.get("nombre")
        and not pedido_actual.get("tipo_entrega")
        and not pedido_actual.get("direccion")
    )

    if mensaje_lower in mensajes_transferencia and pedido_vacio:
        ultimo_pedido = obtener_ultimo_pedido_reciente(numero)

        if (
            ultimo_pedido
            and ultimo_pedido.get("metodo_pago", "").lower() == "transferencia"
        ):
            enviar_texto(
                numero,
                "Claro, te envío nuevamente el código QR:"
            )

            qr_enviado = enviar_imagen(numero)

            if not qr_enviado:
                enviar_texto(
                    numero,
                    "Lo siento, no pude enviar el código QR en este momento."
                )

            return Response(
                content="EVENT_RECEIVED",
                media_type="text/plain"
            )

    if pedido_actual.get("confirmar_direccion_guardada"):
        direccion_guardada = pedido_actual["confirmar_direccion_guardada"]

        respuestas_si = ["si", "sí", "claro", "dale", "ok", "okay", "correcto", "yes"]

        if mensaje_lower in respuestas_si:
            pedido_actual["direccion"] = direccion_guardada
        else:
            pedido_actual["direccion"] = mensaje.strip()

        pedido_actual["confirmar_direccion_guardada"] = ""

        guardar_contacto(
            numero,
            pedido_actual.get("nombre") or None,
            pedido_actual.get("direccion") or None
        )

        guardar_conversacion(numero, pedido_actual)

        enviar_texto(
            numero,
            "Perfecto, ya tengo el número del local para este pedido."
        )

        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    menu_pdf_url = f"{PUBLIC_BASE_URL}/static/menu.pdf"
    qr_url = f"{PUBLIC_BASE_URL}/static/qr_transferencia.jpeg"

    if "menu" in mensaje_lower or "menú" in mensaje_lower:
        enviar_texto(numero, "Claro, aquí tienes nuestro menú 📋")
        enviar_documento(numero, menu_pdf_url, "Menu Terrace.pdf")
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    pedido_actual = obtener_conversacion(numero)

    contacto = obtener_contacto(numero)

    if contacto and contacto.get("nombre"):
        pedido_actual["nombre"] = contacto["nombre"]

    try:
        completion = client.chat.completions.create(
            model="gpt-5",
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

    {json.dumps(MENU, ensure_ascii=False)}

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
    - el método de pago;
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
    "Transferencia" → pedido

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

    Si pide una categoría general que tiene varias opciones, no elijas una arbitrariamente.

    Ejemplos:

    - limonada
    - jugo
    - malteada
    - capuccino
    - sandwich
    - torta

    Agrégala a "productos_ambiguos" y no a "items".

    Si luego aclara la opción, elimina la ambigüedad y agrega el producto exacto.

    PRODUCTOS NO DISPONIBLES:

    Si pide un producto que no existe en el menú:

    - no lo agregues a "items";
    - agrégalo a "productos_no_disponibles";
    - no lo reemplaces por otro producto.

    TIPO DE ENTREGA:

    "tipo_entrega" solo puede ser:

    - "recoger"
    - "domicilio"
    - ""

    Usa "domicilio" para expresiones como:

    - domicilio
    - delivery
    - enviar
    - llévalo
    - me lo traen
    - para el local

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

    MÉTODO DE PAGO:

    "metodo_pago" solo puede ser:

    - "efectivo"
    - "transferencia"
    - ""

    Convierte expresiones como:

    - cash → efectivo
    - QR → transferencia
    - Nequi → transferencia
    - transfiero → transferencia

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
    5. Que "tipo_entrega" y "metodo_pago" tengan valores permitidos.

    IMPORTANTE SOBRE LA INTENCIÓN:

    La intención debe determinarse solamente usando el NUEVO MENSAJE DEL CLIENTE.

    No copies ni conserves el valor anterior de "tipo" o "respuesta".

    Los campos "tipo" y "respuesta" son temporales y describen solamente el mensaje actual.

    Si el nuevo mensaje contiene productos, cantidades, información de entrega,
    dirección, local, nombre, método de pago, notas o correcciones,
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
            response_format={"type": "json_object"}
        )

        data = completion.choices[0].message.content
        print("IA:", data)

        pedido = json.loads(data)

    except Exception as e:
        print("ERROR OPENAI:", e)

        enviar_texto(
            numero,
            "Lo siento, ocurrió un error procesando tu pedido. ¿Podrías intentarlo nuevamente?"
        )

        return Response(
            content="EVENT_RECEIVED",
            media_type="text/plain"
        )

    contacto = obtener_contacto(numero)

    if contacto and contacto.get("nombre") and not pedido.get("nombre"):
        pedido["nombre"] = contacto["nombre"]

    if pedido.get("nombre") or pedido.get("direccion"):
        guardar_contacto(
            numero,
            pedido.get("nombre") or None,
            pedido.get("direccion") or None
        )
    guardar_conversacion(numero, pedido)

    tipo = pedido.get("tipo", "").lower()

    if tipo == "saludo":
        enviar_texto(
            numero,
            "¡Hola! Bienvenido a Terrace ☕🍰\n\n"
            "Te comparto nuestro menú. Cuando estés listo, puedes enviarme tu pedido."
        )
        enviar_documento(numero, menu_pdf_url, "Menu Terrace.pdf")
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    if tipo == "menu":
        enviar_texto(numero, "Claro, aquí tienes nuestro menú 📋")
        enviar_documento(numero, menu_pdf_url, "Menu Terrace.pdf")
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    if tipo == "pregunta":
        enviar_texto(numero, pedido.get("respuesta", ""))
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    faltantes = []
    productos_ambiguos = pedido.get("productos_ambiguos", [])
    productos_no_disponibles = pedido.get("productos_no_disponibles", [])

    tipo_entrega = pedido.get("tipo_entrega", "").strip().lower()

    contacto = obtener_contacto(numero)

    if (
        contacto
        and contacto.get("direccion")
        and tipo_entrega == "domicilio"
        and not pedido.get("direccion")
    ):
        pedido["confirmar_direccion_guardada"] = contacto["direccion"]
        guardar_conversacion(numero, pedido)

        enviar_texto(
            numero,
            f"Ya tenemos registrado este número de local: {contacto['direccion']}.\n\n"
            "¿Deseas usar ese mismo número para este pedido? Responde 'sí' o envía el nuevo número."
        )

        return Response(content="EVENT_RECEIVED", media_type="text/plain")

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
            "Lo siento, no tenemos disponible: " + ", ".join(productos_no_disponibles)
        )
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    if productos_ambiguos:
        mensajes = []

        for ambiguo in productos_ambiguos:
            opciones = opciones_por_producto(ambiguo)

            if opciones:
                mensajes.append(
                    f"¿Cuál opción deseas de {ambiguo}? Tenemos: " + ", ".join(opciones)
                )
            else:
                mensajes.append(
                    f"¿Cuál opción deseas de {ambiguo}?"
                )

        enviar_texto(numero, "\n\n".join(mensajes))
        return Response(content="EVENT_RECEIVED", media_type="text/plain")

    if len(faltantes) == 0:
        guardar_pedido_db(numero, pedido)
        borrar_conversacion(numero)

        items_texto = ", ".join(
            [
                f"{item['cantidad']} x {item['producto']}"
                for item in pedido.get("items", [])
            ]
        )

        total = calcular_total(pedido)

        enviar_texto(
            numero,
            f"Perfecto, recibimos tu pedido: {items_texto}.\n\n"
            f"💰 Total: ${total:,} COP.\n\n"
            "En un momento te confirmamos."
        )

        if pedido.get("metodo_pago", "").strip().lower() == "transferencia":
            enviar_texto(
                numero,
                "Puedes hacer la transferencia usando este código QR:"
            )

            qr_enviado = enviar_imagen(numero)

            if not qr_enviado:
                enviar_texto(
                    numero,
                    "Lo siento, no pude enviar el código QR. "
                    "El personal de Terrace te ayudará con la transferencia."
                )
    else:
        campo = faltantes[0]

        preguntas = {
            "nombre": "Perfecto. ¿A nombre de quien la orden?",
            "items": "¿Qué productos deseas ordenar y en qué cantidades?",
            "tipo_entrega": "¿Es para llevar a un local o pasas a recogerlo?",
            "direccion": "¿Cuál es el numero del local?",
            "metodo_pago": "¿Cuál será el método de pago, tenemos efectivo o transferencia?"
        }

        enviar_texto(numero, preguntas.get(campo, "Me falta información para completar tu pedido."))

    return Response(content="EVENT_RECEIVED", media_type="text/plain")

@app.post("/manual/send")
async def enviar_manual(request: Request):
    form = await request.form()

    telefono = form.get("telefono")
    mensaje = form.get("mensaje")

    if telefono and mensaje:
        enviar_texto(telefono, mensaje)

    return {"success": True}

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

    conn.commit()
    conn.close()

    return {"status": "Chats borrados, contactos guardados"}

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
            telefono,
            MAX(nombre) AS nombre,
            MAX(fecha) AS ultima_fecha,
            COUNT(*) FILTER (WHERE direccion = 'in' AND leido = FALSE) AS no_leidos
        FROM messages
        GROUP BY telefono
        ORDER BY ultima_fecha DESC
    """)

    chats = cursor.fetchall()
    conn.close()

    return templates.TemplateResponse(
        request,
        "inbox.html",
        {"chats": chats}
    )

@app.get("/inbox/chats")
async def inbox_chats():
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)

    cursor.execute("""
        SELECT
            telefono,
            MAX(nombre) AS nombre,
            MAX(fecha) AS ultima_fecha,
            COUNT(*) FILTER (
                WHERE direccion='in'
                AND leido=FALSE
            ) AS no_leidos
        FROM messages
        GROUP BY telefono
        ORDER BY ultima_fecha DESC
    """)

    chats = cursor.fetchall()
    conn.close()

    html = ""

    for chat in chats:

        badge = ""

        if chat["no_leidos"] > 0:
            badge = f"""
            <div class="badge">
                {chat["no_leidos"]}
            </div>
            """

        inicial = (chat["nombre"] or chat["telefono"])[0].upper()

        html += f"""
        <a class="chat" href="/chat/{chat['telefono']}">

            <div class="avatar">
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

    return templates.TemplateResponse(
        request,
        "chat.html",
        {
            "telefono": telefono,
            "mensajes": mensajes
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