select * from campos;

select * from areas;

select * from produc_gas;

select * from planes_prod;

SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'planes_prod';

CREATE VIEW fct_produc_gas AS
select idProduc, fechaProd, ROUND(CAST(VolumenProd AS FLOAT),4) as volumenProduc, idCampo
from produc_gas;

CREATE VIEW fct_planes_prod AS
select idPlan, fechaPlan, ROUND(CAST(volumenPlan AS FLOAT),4) as volumenPlan, idArea
from planes_prod;

CREATE VIEW dim_areas AS
select idArea, NombreArea 
from areas;

CREATE VIEW dim_campos AS
select idCampo, NombreCampo, idArea
from campos;

select * from auditoria_produc_gas;

select * from auditoria_planes_prod;

-- Modificando tabla Areas

UPDATE Areas
SET NombreArea = 'Área ' + CAST(idArea AS VARCHAR)
WHERE NombreArea LIKE 'Area%';
