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
META_APP_ID = os.getenv("META_APP_ID")
META_APP_SECRET = os.getenv("META_APP_SECRET")
META_CONFIG_ID = os.getenv("META_CONFIG_ID", "1200748898865773")
META_GRAPH_VERSION = os.getenv("META_GRAPH_VERSION", "v25.0")

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


def enviar_documento(numero, link, filename="menu.pdf"):
    url = f"https://graph.facebook.com/v25.0/{META_PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "document",
        "document": {
            "link": link,
            "filename": filename
        }
    }

    response = requests.post(url, headers=headers, json=payload)
    print("META DOCUMENT RESPONSE:", response.status_code, response.text)


def enviar_imagen(numero, link):
    url = f"https://graph.facebook.com/v25.0/{META_PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "image",
        "image": {
            "link": link
        }
    }

    response = requests.post(url, headers=headers, json=payload)
    print("META IMAGE RESPONSE:", response.status_code, response.text)

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
    print("META WEBHOOK:", json.dumps(body, ensure_ascii=False))

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

    completion = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": f"""
Eres un agente de pedidos para un negocio de comida.

Tu tarea es actualizar el pedido usando:
1. La información que ya se tenía.
2. El nuevo mensaje del cliente.

Campos:
- nombre
- items
- tipo_entrega
- direccion
- hora
- metodo_pago
- notas

items debe ser una lista así:
[
  {{"producto": "Limonada de coco", "cantidad": "2", "notas": ""}},
  {{"producto": "Malteada de chocolate", "cantidad": "1", "notas": "leche deslactosada"}}
]

Este es el menú disponible:
{json.dumps(MENU, ensure_ascii=False)}

Primero determina la intención del cliente.

Puede ser uno de estos tipos:

- saludo
- pregunta
- menu
- pedido

Si es solamente un saludo, por ejemplo "hola", "buenas", "buenos días":
- responde cordialmente.
- NO inicies un pedido.
- devuelve tipo="saludo".

Si el mensaje incluye saludo + productos, por ejemplo:
"Hola, un latte"
"Buenas, quiero un tinto"
"Buenos días, me das una limonada de coco"
Entonces NO es saludo. Es tipo="pedido" y debes actualizar el pedido.

Si el cliente pide el menú:
- devuelve tipo="menu".
- NO inicies un pedido.

Si el cliente hace una pregunta sobre el restaurante, el menú, productos, horarios, dirección, métodos de pago o cualquier información:
- responde la pregunta.
- NO modifiques el pedido.
- devuelve tipo="pregunta".

Solo cuando el cliente realmente quiera comprar o agregar productos:
- devuelve tipo="pedido".
- actualiza el pedido.

Información del restaurante:

Nombre:
Terrace

Horario:
8:00 AM a 10:00 PM

Dirección:
...

Vendemos:

- Café
- Postres
- Desayunos
- Bebidas
- Almuerzos

Métodos de pago:
Efectivo
Transferencia

Si conoces la respuesta usando esta información, responde naturalmente.

Si no tienes la información, responde que no estás seguro y ofrece ayuda.

Reglas:
- No borres información que ya existe.
- Si el cliente corrige algo, actualízalo.
- La hora y las notas NO son obligatorias.
- El campo tipo_entrega SOLO puede ser "recoger" o "domicilio".
- Si el cliente dice "para llevar", "delivery", "llévalo", "enviar", "me lo traen", conviértelo a "domicilio".
- Si el cliente dice "pickup", "paso por él", "lo recojo", conviértelo a "recoger".
- Si tipo_entrega es "recoger", la direccion puede quedar vacía.
- Si tipo_entrega es "domicilio", la direccion sí es obligatoria.
- Devuelve SOLO JSON válido.
- Solo puedes aceptar productos que estén en el menú.
- Si el cliente pide algo general como "limonada", "jugo", "malteada", "capuccino", "sandwich" o "torta", y existen varias opciones, NO lo agregues a items todavía.
- En ese caso agrégalo a productos_ambiguos.
- Si el cliente pide un producto que no existe en el menú, agrégalo a productos_no_disponibles.
- Si el cliente especifica bien, por ejemplo "limonada de coco", "jugo de mora", "malteada de oreo", sí agrégalo a items.
- Si pide leche deslactosada, agrega en notas "leche deslactosada" para ese producto.
- Si el cliente dice "tinto", eso significa el producto exacto "Tinto".
- Si el cliente dice "un tinto", "1 tinto" o "quiero tinto", agrégalo como:
  {{"producto": "Tinto", "cantidad": "1", "notas": ""}}
"""
            },
            {
                "role": "user",
                "content": f"""
Pedido actual:
{json.dumps(pedido_actual, ensure_ascii=False)}

Nuevo mensaje del cliente:
{mensaje}

Devuelve este JSON:
{{
  "tipo":"",
  "respuesta":"",
  "nombre":"",
  "items":[],
  "productos_ambiguos":[],
  "productos_no_disponibles":[],
  "tipo_entrega":"",
  "direccion":"",
  "hora":"",
  "metodo_pago":"",
  "notas":""
}}
"""
            }
        ]
    )

    data = completion.choices[0].message.content
    print("IA:", data)

    pedido = json.loads(data)
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
            enviar_texto(numero, "Puedes hacer la transferencia usando este código QR:")
            enviar_imagen(numero, qr_url)
    else:
        campo = faltantes[0]

        preguntas = {
            "nombre": "Perfecto. ¿A nombre de quien la orden?",
            "items": "¿Qué productos deseas ordenar y en qué cantidades?",
            "tipo_entrega": "¿Es para recoger o domicilio?",
            "direccion": "¿Cuál es el numero del local?",
            "metodo_pago": "¿Cuál será el método de pago, tenemos efectivo o transferencia?"
        }

        enviar_texto(numero, preguntas.get(campo, "Me falta información para completar tu pedido."))

    return Response(content="EVENT_RECEIVED", media_type="text/plain")



@app.get("/connect-whatsapp", response_class=HTMLResponse)
async def connect_whatsapp_page(request: Request):
    if not META_APP_ID:
        return HTMLResponse(
            "<h2>Falta META_APP_ID en las variables de entorno.</h2>",
            status_code=500,
        )

    return templates.TemplateResponse(
        request=request,
        name="connect_whatsapp.html",
        context={
            "meta_app_id": META_APP_ID,
            "meta_config_id": META_CONFIG_ID,
            "meta_graph_version": META_GRAPH_VERSION,
        },
    )


@app.post("/embedded-signup/callback")
async def embedded_signup_callback(request: Request):
    """Intercambia el código de Embedded Signup por un business token y guarda la cuenta."""
    if not META_APP_ID or not META_APP_SECRET:
        raise HTTPException(
            status_code=500,
            detail="Faltan META_APP_ID y/o META_APP_SECRET en las variables de entorno."
        )

    payload = await request.json()
    code = payload.get("code")
    session = payload.get("session") or {}

    if not code:
        raise HTTPException(status_code=400, detail="No se recibió el código de autorización.")

    token_response = requests.get(
        f"https://graph.facebook.com/{META_GRAPH_VERSION}/oauth/access_token",
        params={
            "client_id": META_APP_ID,
            "client_secret": META_APP_SECRET,
            "code": code,
        },
        timeout=30,
    )

    if not token_response.ok:
        print("ERROR CAMBIANDO CODE:", token_response.status_code, token_response.text)
        raise HTTPException(status_code=400, detail=token_response.json())

    token_data = token_response.json()
    business_token = token_data.get("access_token")
    expires_in = token_data.get("expires_in")

    if not business_token:
        raise HTTPException(status_code=400, detail="Meta no devolvió un access_token.")

    # Embedded Signup puede enviar estos datos dentro de data.
    session_data = session.get("data", session) if isinstance(session, dict) else {}
    waba_id = session_data.get("waba_id")
    phone_number_id = session_data.get("phone_number_id")
    business_id = session_data.get("business_id")

    if not waba_id:
        raise HTTPException(
            status_code=400,
            detail="No se recibió waba_id. Completa el flujo y vuelve a intentarlo."
        )

    # Suscribe esta app a los webhooks de la WABA conectada.
    subscribe_response = requests.post(
        f"https://graph.facebook.com/{META_GRAPH_VERSION}/{waba_id}/subscribed_apps",
        headers={"Authorization": f"Bearer {business_token}"},
        timeout=30,
    )

    if not subscribe_response.ok:
        print("ERROR SUBSCRIBING APP:", subscribe_response.status_code, subscribe_response.text)

    # Si el evento no incluyó phone_number_id, lo consultamos a Meta.
    if not phone_number_id:
        phones_response = requests.get(
            f"https://graph.facebook.com/{META_GRAPH_VERSION}/{waba_id}/phone_numbers",
            headers={"Authorization": f"Bearer {business_token}"},
            timeout=30,
        )
        if phones_response.ok:
            phones = phones_response.json().get("data", [])
            if phones:
                phone_number_id = phones[0].get("id")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO whatsapp_accounts (
            waba_id, phone_number_id, business_id, access_token,
            token_expires_at, connected_at, active
        )
        VALUES (
            %s, %s, %s, %s,
            CASE WHEN %s IS NULL THEN NULL
                 ELSE CURRENT_TIMESTAMP + (%s * INTERVAL '1 second') END,
            CURRENT_TIMESTAMP, TRUE
        )
        ON CONFLICT (waba_id)
        DO UPDATE SET
            phone_number_id = EXCLUDED.phone_number_id,
            business_id = EXCLUDED.business_id,
            access_token = EXCLUDED.access_token,
            token_expires_at = EXCLUDED.token_expires_at,
            connected_at = CURRENT_TIMESTAMP,
            active = TRUE
    """, (
        waba_id,
        phone_number_id,
        business_id,
        business_token,
        expires_in,
        expires_in,
    ))
    conn.commit()
    conn.close()

    return JSONResponse({
        "success": True,
        "waba_id": waba_id,
        "phone_number_id": phone_number_id,
        "business_id": business_id,
        "token_expires_in": expires_in,
        "webhook_subscribed": subscribe_response.ok,
    })

@app.get("/embedded-signup/status")
async def embedded_signup_status():
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute("""
        SELECT waba_id, phone_number_id, business_id, token_expires_at,
               connected_at, active
        FROM whatsapp_accounts
        ORDER BY connected_at DESC
        LIMIT 1
    """)
    account = cursor.fetchone()
    conn.close()
    return {"connected": bool(account), "account": account}

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
        request,
        "dashboard.html",
        {
            "orders": orders_with_items,
            "meta_app_id": META_APP_ID or "",
            "meta_config_id": META_CONFIG_ID
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