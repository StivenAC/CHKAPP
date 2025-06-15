# SIIAPP_CHKAP - Sistema Integral de Información de Aprobaciones

## 1. Purpose and Scope

**SIIAPP_CHKAP** (Sistema Integral de Información de Aprobaciones - Checklist de Aprobaciones) es una plataforma de gestión de aprobaciones tipo checklist para el seguimiento del ciclo de vida de productos, integrando procesos de control de calidad, validación de producción y empaque.

Esta es una aplicación de escritorio desarrollada en **Python** con **CustomTkinter**, que se conecta a bases de datos **SQL Server** y utiliza **autenticación LDAP** contra Active Directory. Está empaquetada como un ejecutable independiente para Windows utilizando **PyInstaller**.

## 2. System Architecture

La arquitectura está separada en capas modulares: presentación, lógica de negocio y acceso a datos.

### Componentes clave:

- **App (ctk.CTk)** – `SIIAPP_CHKAP.py`
- **LoginFrame** – `modules.login_frame`
- **MyTabView** – `modules.tab_view`
- **GUI Frame** – `modules.gui_frame`
- **LDAP Authentication** – `modules.login` (servidor AD: `SERVER2.GBLAB.LOCAL`)
- **Operaciones SQL** – `modules.query`
- **Generación de documentos PDF** – `modules.pdf`
- **Utilidades** – `modules.extra_functions`
- **Formularios** – `modules.f_combobox`

### Bases de datos utilizadas:

- **DB1 (Principal)**: 10.10.10.100
- **DB2 (Administración)**: 10.10.10.10

### Recursos externos:

- `.env` – configuración de entorno
- `credentials.txt` – credenciales cifradas
- `Assets/icon.ico` – ícono de la app
- `themes/lavender.json` – tema UI
- `PDF/Formato_de_aprobación_act.pdf` – plantilla de documento

## 3. Application Structure and Core Components

### Flujo principal:

1. **App.init()** – Inicializa la ventana principal (1000x600)
2. **LoginFrame** – Gestión de autenticación LDAP
3. **show_app_frame(user_access)** – Renderizado de UI basado en permisos

### Jerarquía de componentes:

- `logf.LoginFrame` (login UI)
- `tb.MyTabView` (interfaz con pestañas)
- `tab_configs`, `disabled_tabs`, `active_tab` – Configuración de navegación
- `user_access` – Diccionario con acceso dinámico por usuario

## 4. Technology Stack and Dependencies

| Componente       | Tecnología             | Propósito                                 |
|------------------|------------------------|-------------------------------------------|
| GUI              | customtkinter          | Interfaz gráfica moderna                  |
| Base de datos    | pyodbc                 | Conexión con SQL Server                   |
| Autenticación    | ldap3                  | Autenticación vía Active Directory (LDAP) |
| UI/Grillas       | tksheet                | Componente tipo hoja de cálculo           |
| Configuración    | python-dotenv          | Manejo de variables de entorno            |
| Cifrado          | cryptography.fernet    | Cifrado simétrico                         |
| Empaquetado      | PyInstaller            | Generación de ejecutable                  |
| Runtime          | Python 3.10            | Entorno de ejecución principal            |

## 5. Configuration and Security Architecture

### Archivos de configuración:

- `.env`: configuración de conexiones y dominio
- `credentials.txt`: credenciales cifradas en base64
- `ACCESS_CONFIG`: variables de acceso dinámico

### Seguridad:

- **Conexiones a DB**: autenticación con UID/PWD y cuenta SA
- **Autenticación AD**: `AD_DOMAIN=GBLAB.LOCAL`, `AD_SERVER=SERVER2.GBLAB.LOCAL`
- **Control de acceso**: por grupo (`ALLOWED_GROUPS`) y usuario (`ALLOWED_USERS`)
- **Cifrado**: `ENCRYPTION_KEY` Fernet (ejemplo: `J2XLByvXueRHojClwd5gqin9KZynhzhQuTheo91hnmk=`)

## 6. Application Packaging and Distribution

El sistema es empaquetado como ejecutable de Windows usando PyInstaller. La configuración incluye:

- **Script principal**: `SIIAPP_CHKAP.py`
- **Carpetas incluidas**: `./modules`, `./PDF`
- **Recursos**: `.env`, `credentials.txt`, temas, íconos y plantillas PDF
- **Icono embebido**: `Assets/icon.ico`
- **Imports ocultos**: definidos en `.spec` para detección completa de dependencias

El ejecutable generado incluye el runtime de Python, todas las dependencias y archivos necesarios en un único paquete distribuible.
