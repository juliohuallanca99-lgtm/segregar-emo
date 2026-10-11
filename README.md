# segregar-emo

Separa los PDF de exámenes médicos (EMO) por proyecto, cruzando el nombre del archivo
con el Excel de Fuerza Laboral. Fase actual: solo proyecto **0171 - Movimiento de tierras TSF 3+**.

## Requisitos
    pip install openpyxl

## Uso (primero simulación)
    python segregar_emo.py --lista "P0171_Marcobre_TSF_3_Fuerza_Laboral_2026.xlsx" --emo "C:\ruta\EMO" --destino "C:\ruta\Proyectos Medicos" --meses "1. ENE"

Agregar `--copiar` para copiar de verdad (los originales no se tocan).

## Estados del reporte
- `TSF_EXACTO`: está en la lista, se copia a `0171.../2026/01 ENE/`
- `TSF_REVISAR`: nombre parecido o truncado; NO se copia, revisar a mano
- `NO_TSF`: no está en la lista; no se toca
- `DUPLICADO_IDENTICO` / `TSF_EXACTO_VERSION`: mismo nombre ya en destino

## Datos
No subir PDF ni Excel a este repositorio (datos personales y médicos). Ver `.gitignore`.
