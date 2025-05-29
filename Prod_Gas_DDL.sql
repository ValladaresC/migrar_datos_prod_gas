CREATE DATABASE ProdGas;
GO
USE ProdGas;

CREATE TABLE areas
(
	idArea int primary key,
	NombreArea varchar(20)
);

CREATE TABLE campos 
(
	idCampo int primary key,
	NombreCampo varchar(20),
	idArea int not null,
	constraint fk_AreaCampo foreign key (idArea) references Areas (idArea)
);

CREATE TABLE planes_prod
(
	idPlan int primary key,
	fechaPlan date,
	volumenPlan varchar(20),
	idArea int not null,
	constraint fk_AreaPlan foreign key (idArea) references Areas (idArea)
);

CREATE TABLE produc_gas
(
	idProduc int primary key,
	fechaProd date,
	VolumenProd varchar(20),
	idCampo int not null, 
	constraint fk_CampoProduc foreign key (idCampo) references campos (idCampo)
);

DROP TABLE produc_gas;
DROP TABLE planes_prod;
DROP TABLE campos;
DROP TABLE areas;

TRUNCATE TABLE produc_gas;
TRUNCATE TABLE planes_prod;
TRUNCATE TABLE campos;
TRUNCATE TABLE areas;


CREATE TABLE auditoria_produc_gas (
    auditoriaID INT IDENTITY(1,1) PRIMARY KEY,
    operacion VARCHAR(10), -- 'INSERT', 'UPDATE', 'DELETE'
    fecha DATETIME DEFAULT GETDATE(),
    datosAntes NVARCHAR(MAX),
    datosDespues NVARCHAR(MAX)
);

CREATE TRIGGER trg_auditar_produc_gas_insert
ON produc_gas
AFTER INSERT
AS
BEGIN
    INSERT INTO auditoria_produc_gas (operacion, datosDespues)
    SELECT
        'INSERT',
        (SELECT * FROM INSERTED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;

CREATE TRIGGER trg_auditar_produc_gas_update
ON produc_gas
AFTER UPDATE
AS
BEGIN
    INSERT INTO auditoria_produc_gas (operacion, datosAntes, datosDespues)
    SELECT
        'UPDATE',
        (SELECT * FROM DELETED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER),
        (SELECT * FROM INSERTED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;

CREATE TRIGGER trg_auditar_produc_gas_delete
ON produc_gas
AFTER DELETE
AS
BEGIN
    INSERT INTO auditoria_produc_gas (operacion, datosAntes)
    SELECT
        'DELETE',
        (SELECT * FROM DELETED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;

CREATE TABLE auditoria_planes_prod (
    auditoriaID INT IDENTITY(1,1) PRIMARY KEY,
    operacion VARCHAR(10), -- 'INSERT', 'UPDATE', 'DELETE'
    fecha DATETIME DEFAULT GETDATE(),
    datosAntes NVARCHAR(MAX),
    datosDespues NVARCHAR(MAX)
);

CREATE TRIGGER trg_auditar_planes_prod_insert
ON planes_prod
AFTER INSERT
AS
BEGIN
    INSERT INTO auditoria_planes_prod (operacion, datosDespues)
    SELECT
        'INSERT',
        (SELECT * FROM INSERTED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;

CREATE TRIGGER trg_auditar_planes_prod_update
ON planes_prod
AFTER UPDATE
AS
BEGIN
    INSERT INTO auditoria_planes_prod (operacion, datosAntes, datosDespues)
    SELECT
        'UPDATE',
        (SELECT * FROM DELETED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER),
        (SELECT * FROM INSERTED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;

CREATE TRIGGER trg_auditar_planes_prod_delete
ON planes_prod
AFTER DELETE
AS
BEGIN
    INSERT INTO auditoria_planes_prod (operacion, datosAntes)
    SELECT
        'DELETE',
        (SELECT * FROM DELETED FOR JSON PATH, WITHOUT_ARRAY_WRAPPER)
END;