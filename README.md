# Hand in Hand

**Estamos aquí para darte una mano.**

Agencia digital que conecta micro y pequeños emprendimientos con nano influencers verificados y demuestra el retorno de cada campaña con códigos y enlaces únicos.

> Enlace público: [equipo-quipitos.github.io/CLASE-](https://equipo-quipitos.github.io/CLASE-/)
>
> Hoja de cálculo (base de datos del MVP): [ver hoja](https://docs.google.com/spreadsheets/d/1Y43if2kr41J8EeVmhEKu_hNu2dDx7uDJ3fBbrbph4cM/edit?usp=sharing)

---

## 1. Problema

Los propietarios de micro y pequeños emprendimientos enfrentan una desventaja competitiva debido a su dificultad para llegar a su público con credibilidad en canales digitales. Aunque reconocen la necesidad de apoyarse en creadores de contenido para promocionarse, la imposibilidad económica de acceder a agencias tradicionales, la alta carga de trabajo operativo y la falta de herramientas para medir resultados les impiden gestionar estas alianzas de forma efectiva, limitando la captación de nuevos clientes y su crecimiento comercial.

### Necesidades
- Atraer nuevos clientes y generar ventas para sostener y hacer crecer el negocio.
- Ganar credibilidad y reconocimiento frente a su público.

### Dificultades
- Las agencias y plataformas de influencers están pensadas para grandes presupuestos y no atienden a los emprendimientos pequeños.
- Buscar, verificar y gestionar creadores por su cuenta consume un tiempo que no tienen.
- No cuentan con herramientas para saber si lo que invirtieron generó clientes.

### Expectativas
- Llegar a su público a través de voces creíbles.
- Hacerlo sin agencias costosas y sin dedicarle demasiado tiempo.
- Saber con claridad qué resultados obtuvo cada inversión.

## 2. Insight

Lo que las grandes agencias descartan por pequeño es justo lo que da confianza y permite medir con certeza. Los nano influencers tienen audiencias reducidas pero cercanas y creíbles, y precisamente por ser pequeñas resultan más fáciles de rastrear con códigos y enlaces únicos. Lo que parecía una limitación de alcance es una ventaja de precisión.

## 3. Concepto de solución

Una agencia digital, apoyada en una plataforma web puente, que conecta a emprendimientos con nano influencers curados y les demuestra el retorno de cada campaña mediante códigos únicos, enlaces rastreables y un tablero de resultados. Se vende certeza de retorno, no solo alcance.

| Característica | Atributo |
|---|---|
| Confianza y credibilidad | Perfiles verificados de nano influencers |
| Certeza de retorno | Códigos únicos, enlaces UTM y tablero de métricas |
| Sencillez | Registro con formulario en pocos pasos |
| Modelo accesible | Comisión del 20 al 30 % por campaña gestionada |

## 4. MVP

Se valida primero de forma manual, sin construir la plataforma completa.

| Componente | Cómo funciona |
|---|---|
| Landing page | Explica la propuesta con la identidad visual Hand in Hand, con un tutorial guiado (`index.html`) |
| Registro | Una página con dos pestañas: "Soy emprendedor" (verde) y "Soy un creador" (azul), con listas desplegables (`registro.html`) |
| Base de datos | Los registros llegan a una hoja de cálculo de Google mediante Google Apps Script (`Code.gs`) |
| Correo de bienvenida | Cada persona registrada recibe un correo de agradecimiento con los colores de la página, y el equipo recibe un aviso |
| Curaduría | 15 a 20 nano influencers verificados manualmente, de 1 o 2 sectores |
| Emparejamiento | Manual: el equipo cruza cada emprendimiento con creadores afines y coordina por WhatsApp |
| Trazabilidad | Código único y enlace UTM por creador, con tablero en Sheets + Looker Studio |
| Piloto | 3 a 5 emprendimientos, una campaña cada uno, durante 4 semanas |

### Registro y base de datos

El piloto se enfoca en el departamento del **Magdalena**, empezando por Santa Marta. Por eso el municipio, el sector y el nicho se eligen de listas desplegables, para que los datos lleguen escritos siempre igual y sea fácil cruzar negocios y creadores del mismo lugar.

Cada formulario guarda una fila en su pestaña de la hoja. El script ubica cada dato según los títulos de la fila 1, y llena solo `id`, `fecha`, `estado` y `verificado`.

| Pestaña | Columnas |
|---|---|
| Emprendimientos | `id`, `fecha`, `negocio`, `sector`, `ciudad` (municipio), `presupuesto`, `objetivo`, `tiempo_disponible`, `url_destino`, `whatsapp`, `correo`, `estado` |
| Creadores | `id`, `fecha`, `nombre`, `red`, `usuario_red`, `seguidores`, `nicho`, `ciudad` (municipio), `tarifa`, `whatsapp`, `correo`, `verificado`, `estado` |

### Hipótesis a validar
1. Los emprendimientos pequeños están dispuestos a pagar por campañas con retorno medible.
2. Los nano influencers aceptan trabajar con negocios pequeños bajo estos términos.
3. La trazabilidad por códigos y enlaces permite atribuir ventas a cada creador.

### Métricas de éxito (metas de partida, a ajustar)
- Más del 50 % de las propuestas se concreta en colaboración.
- Cada campaña muestra ventas o consultas atribuidas por código.
- Más del 70 % de los emprendimientos repetiría.
- Al menos la mitad acepta el esquema de comisión.

### Riesgo principal
Con presupuestos pequeños, una comisión del 20 al 30 % deja poco ingreso por campaña. En el piloto se probará también una tarifa mínima o paquetes de varias campañas.

## 5. Estructura del proyecto

```
├── assets/
│   └── config.js    configuración (API_URL del Apps Script)
├── index.html       página de inicio, con sus estilos y el tutorial guiado
├── registro.html    formularios de emprendimientos y creadores, con sus estilos y el envío de datos
├── Code.gs          copia del Google Apps Script (guarda registros y envía el correo de bienvenida)
└── README.md        este documento
```

## 6. Tecnologías

- **HTML**: estructura de las páginas.
- **CSS**: colores (paleta Hand in Hand), tipografías Inter y Dancing Script, y diseño adaptable a celular. Va dentro de cada archivo HTML.
- **JavaScript**: tutorial guiado, pestañas del registro y envío de formularios. Va dentro de cada archivo HTML.
- **Google Apps Script**: recibe los formularios, los guarda en la hoja y envía los correos.
- **Google Sheets**: base de datos de registros y seguimiento de campañas.
- **GitHub Pages**: publicación del sitio.

## Equipo

Equipo Quipitos
