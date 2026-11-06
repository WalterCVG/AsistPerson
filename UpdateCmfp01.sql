'UPDATE A
'    SET A.codniv = B.CodNiv, A.fecnac=B.fecnac, A.documento=B.documento   
'FROM cmfp01 A
'    INNER JOIN Cpgeneral B ON 
'    	A.tarjeta = B.tarjeta

'UPDATE    Histo
'SET       DNI = t2.DNI
'FROM      TABLADNI AS t2 
'WHERE     t2.HISTORIA = Histo.HISTORIA


'UPDATE A
'    SET A.cargo = B.Codi   
'FROM CpGeneral A
'    INNER JOIN dat_nomb B ON 
'    	A.documento = B.libele and a.cargo<>b.codi


'-Copiar sólo la estructura de Tabla :
'SELECT * Into DestinationTableName From SourceTableName Where 1 = 2 

'- Duplicado exacto de una Tabla :
'SELECT * INTO MyNewTable FROM MyTable


'-Reemplazar caracteres

'UPDATE tabla
'SET nombre = Replace(nombre, 'ñ', 'Ñ')

'select dbo.CHRTRAN('ABCDEF', 'ACE', 'XYZ') -- Muestra 'XBYDZF'
'select dbo.CHRTRAN('ABCDEF', 'ACE', 'XYZQRST') -- Muestra 'XBYDZF'



'- insertar registros entre 2 tablas cuya estructura es igual
'insert into tabla-B  select * from tabla-A

'- insertar registros entre 2 tablas cuya estructura es diferente, solo indico los campos iguales de tabla destino
'Insert Into tabla-B (campo1, campo2, campo3, campo4, campo5)  select * from tabla-A

'-OK
'INSERT INTO CpAutori2012
'                      (Tarjeta, CODIGO, DiaHora, FECHAINI, FECHAFIN, DOCUMENTO, USUARIO, FECHA)
'SELECT     Tarjeta, CODIGO, DiaHora, FECHAINI, FECHAFIN, DOCUMENTO, USUARIO, FECHA
'FROM         CpAuPrg2012

'WHERE --<Condicion> 


'-Consulta cruzada
'SELECT   tb1.*, tb2.Tarjeta AS codigo, tb2.Plh AS hepcas, tb2.TipoPla AS planilla
'FROM     CpMarca2012 tb1 INNER JOIN
'         CpGeneral tb2 ON tb2.Tarjeta = tb1.Tarjeta AND tb2.Plh = 'HEP' AND (tb2.TipoPla = '1' OR
'         tb2.TipoPla = '2')
'ORDER BY tb1.Tarjeta, tb1.Fecha, tb1.HORA

'++++++++++
SELECT     t1.HISTORIA, t1.DNI, t2.HISTORIA AS hc, t2.DNI AS libele
FROM       Histo AS t1 INNER JOIN
                      TABLADNI AS t2 ON t2.HISTORIA = t1.HISTORIA
+++++++++++++++++++++++++++++

'SELECT   tb1.*, tb2.Tarjeta AS codigo, tb2.Plh AS hepcas, tb2.TipoPla AS planilla
'FROM     CpDirectorio tb1 INNER JOIN
'         CpGeneral tb2 ON tb2.Tarjeta = tb1.Tarjeta AND tb1.ANO = '2012' AND tb1.MES = '10' AND tb2.Plh = 'HEP' AND (tb2.TipoPla = '1' OR
'         tb2.TipoPla = '2')
'ORDER BY tb1.Tarjeta, tb1.ANO, tb1.MES


==================================

MODIFICAR ESTRUCTURA DE TABLA
=================================

'-SQL Server no permite agregar campos "not null" a menos que se especifique un valor por defecto:

 alter table libros
  add autor varchar(20) not null default 'Desconocido';


'-Para eliminar campos de una tabla la sintaxis básica es la siguiente:

 alter table NOMBRETABLA
  drop column NOMBRECAMPO;

'-No pueden eliminarse los campos que son usados por un índice o tengan restricciones. No puede eliminarse un campo si es el único en la tabla.

'-Podemos eliminar varios campos en una sola sentencia:

 alter table libros
  drop column editorial,edicion;

'-En la tabla Customer, deseamos renombrar “Address" a “Addr”.  Para hacerlo, ingresamos,

ALTER TABLE Customer
 CHANGE Address Addr char(50);