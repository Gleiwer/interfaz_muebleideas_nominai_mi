
# coding: utf-8

# In[ ]:

import pandas as pd
import os
import configparser
import datetime as dt     


# In[ ]:

config = configparser.ConfigParser()
config.read("./config.ini")
in_path=config.get("rutas", "in")
out_path=config.get("rutas","out")

def ConfigSectionMap(section):
    dict1 = {}
    options = config.options(section)
    for option in options:
        try:
            dict1[option] = config.get(section, option)
            if dict1[option] == -1:
                DebugPrint("skip: %s" % option)
        except:
            print("exception on %s!" % option)
            dict1[option] = None
    return dict1


# In[ ]:



for file in os.listdir(in_path):
    if file.endswith(".xlsx"):
        file=os.path.join(in_path, file)
        in_file = pd.read_excel(file)


# In[ ]:

del in_file["Codigo"]
del in_file["Empleado"]
del in_file["VTOTAL"]


# In[ ]:

concept_dict=ConfigSectionMap("conceptos")
column_names=[concept_dict["rn"],
              concept_dict["he"],
              concept_dict["hen"],
              concept_dict["hfd"],
              concept_dict["hefn"],
             "CC"]
in_file.columns=column_names+['Cedula']
in_file.set_index(['Cedula','CC'],inplace=True)
in_file_stacked=in_file.stack()


# In[ ]:

in_file


# In[ ]:

stacked_df=pd.DataFrame(in_file_stacked)
stacked_df.reset_index(level=[0,1,2], inplace=True)


# In[ ]:

stacked_df


# In[ ]:

stacked_df=stacked_df[stacked_df[0]!=0]
out_frame=pd.DataFrame(columns=["Cedula","Codigo Concepto","Codigo Centro de costos",
                   "Codigo Concepto Referencia","Horas", "Valor", "Periodo", 
                    "Fecha", "Salario", "Unidades Producidas", "Es Prestacion", 
                    "Numero Prestamo",	"Dias mes 1", "Dias mes 2",	
                    "Fecha Inicio",	"Base"])


# In[ ]:

out_frame[["Cedula","Codigo Centro de costos","Codigo Concepto","Horas"]]=stacked_df
out_frame["Fecha"]=dt.datetime.today().strftime("%m/%d/%Y")
out_frame["Periodo"]=dt.datetime.today().strftime("%m")
out_frame["Valor"]=0
out_frame["Codigo Concepto Referencia"]=""
out_frame["Salario"]=0
out_frame["Unidades Producidas"]=0
out_frame["Es Prestacion"]=""
out_frame["Numero Prestamo"]=""
out_frame["Dias mes 1"]=0
out_frame["Dias mes 2"]=0
out_frame["Fecha Inicio"]=""
out_frame["Base"]=0
out_frame["Horas"]=out_frame["Horas"]


# In[ ]:

out_frame.set_index("Cedula",inplace=True)


# In[ ]:

out_frame.to_excel(out_path+"/novedades_formato_nomina.xlsx")

