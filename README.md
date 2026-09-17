# AVICO — Gemelo Digital Inteligente de Ginecología y Obstetricia

V1 de un gemelo digital hospitalario para simulación operativa, flujo de pacientes, utilización de recursos y exploración de capacidad.

## V1
- Monitor operativo.
- Modelo PHEDS: 45 posiciones, 27 censables y 18 funcionales/no censables.
- Flujo Triage → Tococirugía/Expulsión → Recuperación → Piso → Egreso.
- Simulación reproducible de 7 días.
- Predict +2/+4/+6 h (prototipo operativo, no clínicamente validado).
- Decision Lab para escenarios contrafactuales.
- Personal DEMO anonimizado.
- Registro de fuentes y parámetros de calibración.
- Histórico local de la última ejecución.

## Despliegue en Render
El repositorio incluye `render.yaml`. En Render seleccione **New + → Blueprint**, conecte este repositorio y aplique el Blueprint.

## Seguridad y alcance
No incluir PII en el repositorio. Los parámetros V1 son de calibración/simulación y no constituyen recomendaciones clínicas ni un sistema clínico validado.

AVICO® · Ginecología y Obstetricia
