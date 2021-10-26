# -*- coding: utf-8 -*-

import pandas as pd
from pandas import ExcelWriter

#Abro excel con la base de datos
df=pd.read_excel("BD.xlsx")
#Ordeno columnas del excel
OrdenColumnas= pd.DataFrame(df, columns = ["Nombre", "Apellido","Departamento","Sección","Matrícula","Salario","Fecha ingreso"])
#Filtro Copiadoras
Copiadoras=df[df.Sección == "Copiadoras"]
FOcopiadoras=pd.DataFrame(Copiadoras, columns = ["Nombre","Apellido","Departamento","Sección","Matrícula","Salario","Fecha ingreso"])
#Filtro Impresoras
Impresoras=df[df.Sección == "Impresoras"]
#Filtro Fax
Fax=df[df.Sección == "Fax"]

#Suma total de salario
SumaSalario=df.Salario.sum()

#Genero un excel con el resultado
writer = ExcelWriter("OrdenColumnas.xlsx")
OrdenColumnas.to_excel(writer,"ETL", index=False)
writer.save()

writerFiltro=ExcelWriter("FiltroCopiadoras.xlsx")
FOcopiadoras.to_excel(writerFiltro, "ETL", index=False)
writerFiltro.save()