-- Drop table

-- DROP TABLE db_ibcms_sql.dbo.Personas;

CREATE TABLE db_ibcms_sql.dbo.Personas
(
    nombres          varchar(100) COLLATE Modern_Spanish_CI_AS NULL,
    apellidos        varchar(100) COLLATE Modern_Spanish_CI_AS NULL,
    fecha_nacimiento date                                      NULL,
    salario          decimal(10, 2)                            NULL,
    id               bigint                                    NOT NULL,
    CONSTRAINT Personas_PK PRIMARY KEY (id)
);
