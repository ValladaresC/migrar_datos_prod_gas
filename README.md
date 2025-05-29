# migrar_datos_prod_gas

> [!WARNING]
> Por seguridad de los datos y políticas de confindecialidad de la empresa, los dataset quedan excluidos de los archivos adjuntos.

El proyecto se centró en abordar la vulnerabilidad en la gestión de datos de producción de gas, los cuales actualmente se almacenan exclusivamente en archivos Excel desde 2021. Esta dependencia presenta riesgos de pérdida de información debido a posibles errores, fallos en los archivos, corrupción o limitaciones en la gestión de grandes volúmenes de datos. 

Debido a que estos datos son fundamentales para otros departamentos, ya que sustentan estadísticas y la planificación estratégica de la alta dirección, su protección y accesibilidad son prioritarias, por lo tanto para mitigar los riesgos, se propuso migrar toda la base de datos a un sistema más robusto y confiable, como MS SQL Server, que además es de código abierto y gratuito. La estrategia incluyó una migración masiva de todos los datos históricos mediante macros que conectan y transfieren la información de Excel hacia SQL Server, asegurando la integridad y respaldo de los datos. 

Posteriormente, se desarrolló un cuadro de mando con botones y macros que permiten la inserción y actualización diaria de la producción de gas, facilitando una gestión eficiente y en tiempo real de la información, mejorando la seguridad, disponibilidad y control de los datos para la toma de decisiones estratégicas.

Tras la migración se logró reducir al 100% la pérdida de datos e incrementar la integridad y fiabilidad de los datos, pasando de 6 incidentes semestrales a 0 incidentes post-implementación. Adicional la disponibilidad y seguridad de la información, lográndose una disminución del 75% de los tiempos de procesamiento para la generación de informes estratégicos.

### 1. Revisión de la base de datos de producción de gas contenida en Excel (Origen)
En esta etapa se identificaron las columnas o campos presentes en el archivo llamado DATAGAS.xlsx (del cual solo se muestra la siguiente imagen), el cual tiene la fecha del registro, la producción por campo en diferentes columnas, la sumatoria de la producción por área y los planes de producción por área. Los valores numéricos se encuentran en formato de miles, es decir 28,481 que es lo mismo que 28481.00.

![image](https://github.com/user-attachments/assets/119284f7-2819-497b-b6ef-1163210a0979)

### 2. Modelamiento de datos.
Se llevo a cabo en 3 etapas:

#### 2.1. Modelo Entidad-Relación:
Se identificaron las principales entidades del sistema, como Áreas, Campos, Planes de Producción y Producción de Gas. Se definieron las relaciones entre ellas:
- Cada Campo pertenece a un Área (relación uno a muchos).
- Cada Plan de Producción está asociado a un Área.
- Cada registro de Producción de Gas está vinculado a un Campo.

#### 2.2. Modelo Lógico: 
Se tradujo el diagrama E-R en tablas relacionales, definiendo las claves primarias y foráneas para asegurar la integridad referencial. Por ejemplo:
- La tabla areas con su clave primaria idArea.
- La tabla campos con su clave primaria idCampo y una clave foránea que referencia a areas.
- La tabla planes_prod con su clave primaria idPlan y una clave foránea a areas.
- La tabla produc_gas con su clave primaria idProduc y una clave foránea a campos.

Se definieron también los tipos de datos apropiados para cada campo.

> [!IMPORTANT]
> La carga de valores decimales en los motores de bases de datos como SQL Server desde Excel (Latinoamérica) generan incompatibilidad entre los tipos de datos, este problema se debe generalmente al uso de los separadores de miles y decimales, diferentes en ambos softwares. Por tal motivo los valores numéricos como los planes de producción y volúmenes de producción se establecieron como varchar.

#### 2.3. Modelo Físico:
Se implementó el modelo lógico en SQL Server creando las tablas con las instrucciones CREATE TABLE, incluyendo restricciones de clave primaria y foránea. Se especificaron los tipos de datos (como int, varchar, date) y las relaciones de integridad referencial para garantizar la consistencia de los datos en la base. Ver el archivo Prod_Gas_DDL.sql

#### 2.4. Agregación de auditoría:
Se crearon tablas de auditoría (auditoria_produc_gas y auditoria_planes_prod) para las tablas más importantes del modelo que son produc_gas y planes_prod. Para alimentar estas tablas se utilizaron Triggers o disparadores, los cuales ante cualquier operación como insert, update o delete, guardarán la fecha/hora, datos y tipo de operación ejecutada. Ver el archivo Prod_Gas_DDL.sql

#### 2.5. Creación de maqueta del modelo dimensional:
Fue creado este modelo, aunque en esta etapa del proyecto aun no se tenía pensado su utilización.

![image](https://github.com/user-attachments/assets/d8709c8b-f5f3-4cbe-b7d1-5fd133e915e5)

### 3. Proceso ETL.
Se creo un archivo Excel nuevo llamado GO_GAS (habilitado para macros) donde se lleva a cabo el proceso ETL.

#### 3.1 Extracción:
Mediante la herramienta Obtener Datos de Excel se realizó el llamado al archivo origen DATAGAS.xlsx

#### 3.2 Transformación:
Mediante la herramienta Power Query se hace una separación del contenido origen en diferentes tablas, quedando las mismas en diferentes hojas: ProducGas, PlanesProd, Campos, Areas; todas ellas con la misma estructura de las tablas destino en SQL Server.

> [!IMPORTANT]
> Se formatearon las columnas de los volúmenes de planes y producción, llevando los valores decimales de coma hacia punto y convirtiendo estas columnas en tipo texto para que al hacer la carga coincida con la estructura tipo varchar de SQL Server. 

En este mismo proceso se adicionó una hoja llamada Menú, la cual actúa como interfase para la interacción del administrador encargado de hacer la carga de datos, la cual tiene dos bloques de Menú: Inserción Total e Inserción Diaria.

![image](https://github.com/user-attachments/assets/e05bb2e9-2f15-408c-9653-9e5ff880e681)

#### 3.3 Carga:
En esta etapa fueron realizadas las distintas macros que dan funcionalidad al Menú y que realizan la carga de datos hacia SQL Server. Las macros de inserción total sirvieron para hacer una carga total de los datos desde 2021 hasta la actualidad. Y las macros de inserción diaria sirvieron y servirán para realizar la migración diaria, semanal o mensual de la data, así también permite actualizar cualquier valor ya cargado. 

Los principales retos de estas macros fueron:
- La compatibilidad en los formatos de las columnas de origen y destino. Uno de estos retos fue nombrado en los Warnings anteriores. Otro reto fue formatear la fecha en la macro para hacerla coincidir con el formato más común usado en SQL Server (YYYY-MM-DD).
- La inserción de miles de datos ya que en una sola operación puede ser muy lento o generar errores si no se hace de manera óptima. La macro implementó un sistema de batching (por ejemplo, cada 500 filas) para mejorar el rendimiento y reducir la carga en la base de datos destino. 
- La construcción segura de cadenas SQL, ya que concatenar grandes cadenas de valores para las sentencias INSERT o UPDATE puede ser complejo y propenso a errores, especialmente si los datos contienen comillas o caracteres especiales. La macro debe asegurarse de formatear correctamente los datos y manejar casos especiales para evitar errores de sintaxis.
- Todas las macros pueden ser vista en la carpeta Macros de los archivos adjuntos o en el archivo GO_GAS.xlsm.

### 4. Creación de visualizaciones en Power BI

Tomando como referencia los informes presentados en el pasado a la junta directiva y como última actividad ejecutada, se creó un dashboard con Power Bi que permitió a la junta directiva de la empresa observar de forma más ordenada y clara el comportamiento de los datos en sus reuniones semanales y mensuales para la toma de decisiones estratégicas.

Este nuevo informe tiene la ventaja de actualizarse luego de haber realizado la carga de datos en SQL Server ya que su modelo dimensional es alimentado mediante vistas SQL. Observar vistas creadas en el archivo Prod_Gas_DML.sql

Su modelo dimensional quedó estructurado de la siguiente manera, y se observa una tabla organizadora de medidas DAX. Consta de:
- Dos tablas dimensionales provenientes del modelo relacional, dim_areas y dim_campos
- Una tabla dimensional creada con lenguaje DAX para la gestión de fechas, dim_calendar
- Dos tablas de hechos provenientes del modelo relacional, fct_planes_prod y fct_produc_gas

![image](https://github.com/user-attachments/assets/2a6b2753-ce02-4448-9b0d-639961861421)

Todas la tablas provenientes del modelo relacional fueron vinculadas con Power BI mediante la creación de vistas SQL y las cuales se actualizan automáticamente al abrir el ejecutable de Power Bi Desktop o solo con presionar el botón 'Actualizar'.

El dashboard consta de:
- Portada: Con botones que dirigen al usuario a la hoja de su interés.
- 2 hojas donde se muestra la variación semanal de producción total, por área y campos, con sus respectivos filtros por año, semana y área.
- 2 hojas donde se muestra la variación mensual de producción total, por área y campos, con sus respectivos filtros por año, semana y área.

![image](https://github.com/user-attachments/assets/2c656fc6-2b1b-4d65-8d2c-010c764a4d57)



