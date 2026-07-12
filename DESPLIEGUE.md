# Desplegar el dashboard en la web (gratis)

Stack: **GitHub + Streamlit Community Cloud + Turso**. Los datos viven en Turso
(nube), así que **no necesitas el PC prendido**: entras desde cualquier
navegador (PC o celular) y todo se gestiona desde el propio dashboard.

La app tiene dos modos automáticos:

- **Modo local** (sin secretos): usa `users/<usuario>/finanzas.db` y Excel,
  igual que siempre. `main.py` y `ejecutar_app.bat` siguen funcionando.
- **Modo nube** (con secretos `[usuarios]`): login con contraseña y datos en
  Turso. Los Excel se generan bajo demanda con el botón "Exportar a Excel".

---

## Paso 1 — Crear las bases de datos en Turso

En [app.turso.tech](https://app.turso.tech) (tu cuenta ya creada):

1. **Create Database** → nombre: `finanzas-juan-manuel` (elige la región más
   cercana, p. ej. São Paulo).
2. Repite para `finanzas-julieta`.
3. Para **cada** base de datos, guarda dos cosas:
   - La **URL** (`libsql://finanzas-xxx-TUORG.turso.io`).
   - Un **token**: en la página de la base de datos → *Create Token*
     (permisos de lectura/escritura, sin expiración o la que prefieras).

## Paso 2 — Migrar tus datos actuales a Turso

Desde la carpeta del proyecto, con el entorno conda `finanzas` activo:

```bat
python migrar_a_turso.py --usuario "juan manuel" --url libsql://finanzas-juan-manuel-XXX.turso.io --token TOKEN_DE_JUAN
python migrar_a_turso.py --usuario "julieta" --url libsql://finanzas-julieta-XXX.turso.io --token TOKEN_DE_JULIETA
```

El script abre la base local en **solo lectura** (no modifica nada local) y
exige que la base de Turso esté vacía (no duplica datos).

## Paso 3 — Subir el código a GitHub

El repo remoto ya existe: `https://github.com/Taisaku32/ANT_FINANCE`.
Verifica en GitHub → Settings que sea **privado** (recomendado). El
`.gitignore` ya excluye bases de datos, Excel, carpeta `users/` y secretos:
**ningún dato personal sube a GitHub**, solo código.

```bat
git add -A
git commit -m "Modo nube: Turso + login + registro desde el dashboard"
git push origin master
```

## Paso 4 — Desplegar en Streamlit Community Cloud

En [share.streamlit.io](https://share.streamlit.io):

1. **New app** → conecta tu GitHub y elige el repo `Taisaku32/ANT_FINANCE`,
   rama `master`, archivo principal `dashboard.py`.
2. Antes de desplegar (o después, en *Settings → Secrets*), pega los secretos
   con el formato de [secrets.ejemplo.toml](secrets.ejemplo.toml), poniendo
   las URLs y tokens reales de Turso y las contraseñas que quieras para
   cada usuario.
3. **Deploy**. Te dará una URL tipo `https://tuapp.streamlit.app`.

Desde el celular: abre esa URL, inicia sesión y listo. Puedes "instalarla"
como app con *Agregar a pantalla de inicio* en el navegador.

## Probar el modo nube desde tu PC (opcional)

Crea `.streamlit/secrets.toml` (está en `.gitignore`, nunca se sube) con el
mismo contenido de los secretos y ejecuta `streamlit run dashboard.py`:
la app local se conectará a Turso igual que la versión web.
**Ojo:** guarda ese archivo en UTF-8 **sin BOM** (en VS Code: "UTF-8", no
"UTF-8 with BOM"), o Streamlit no podrá leerlo.

## Notas

- Streamlit Cloud "duerme" la app tras días sin uso; el primer acceso después
  tarda ~1 minuto en despertarla. Los datos nunca se pierden (están en Turso).
- Los planes gratuitos de Turso (9 GB) y Streamlit sobran para uso personal.
- Si un día quieres volver a local: los archivos de `users/` siguen intactos.
