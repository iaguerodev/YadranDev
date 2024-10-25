import win32com.client
import time

def main():
    try:
        # Conectar con SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Diccionario para asignar códigos según el tipo de gelpack
        gelpack_codes = {
            "gelpack 250 GR": "008",
            "gelpack 500 GR": "009",
            "gelpack 750 GR": "010",
            "gelpack 1000 KG": "011",
            "gelpack 2000 KG": "015"
        }

        # Diccionario para asignar gelpack según el código de material
        material_gelpack = {
            "12000882": "gelpack 500 GR",
            "12000881": "gelpack 500 GR",
            "12000678": "gelpack 500 GR",
            "12000679": "gelpack 500 GR",
            "12000680": "gelpack 500 GR",
            "12000681": "gelpack 500 GR",
            "12000396": "gelpack 500 GR",
            "12000201": "gelpack 500 GR",
            "12000202": "gelpack 500 GR",
            "12000314": "gelpack 500 GR",
            "12000187": "gelpack 500 GR",
            "12000192": "gelpack 500 GR",
            "12000186": "gelpack 500 GR",
            "12000877": "gelpack 500 GR",
            "12000200": "gelpack 500 GR",
            "12000199": "gelpack 500 GR",
            "12000309": "gelpack 500 GR",
            "12000915": "gelpack 500 GR",
            "12000307": "gelpack 500 GR",
            "12000330": "gelpack 500 GR",
            "12000901": "gelpack 500 GR",
            "12000755": "gelpack 500 GR",
            "12000756": "gelpack 500 GR",
            "12000754": "gelpack 500 GR",
            "12000900": "gelpack 500 GR",
            "12000289": "gelpack 500 GR",
            "12000290": "gelpack 500 GR",
            "12000291": "gelpack 500 GR",
            "12000904": "gelpack 500 GR",
            "12000373": "gelpack 500 GR",
            "12001125": "gelpack 500 GR",
            "12001158": "gelpack 500 GR",
            "12001159": "gelpack 500 GR",
            "12000970": "gelpack 500 GR",
            "12000971": "gelpack 500 GR",
            "12000973": "gelpack 500 GR",
            "12000974": "gelpack 500 GR",
            "12000265": "gelpack 500 GR",
            "12000210": "gelpack 250 GR",
            "12000966": "gelpack 750 GR",
            "12000967": "gelpack 750 GR",
            "12000968": "gelpack 750 GR",
            "12000969": "gelpack 750 GR",
            "12000975": "gelpack 750 GR",
            "12000976": "gelpack 750 GR",
            "12000977": "gelpack 750 GR",
            "12000312": "gelpack 750 GR",
            "12000737": "gelpack 750 GR",
            "12000193": "gelpack 750 GR",
            "12000207": "gelpack 750 GR",
            "12000279": "gelpack 750 GR",
            "12000188": "gelpack 750 GR",
            "12000184": "gelpack 750 GR",
            "12000189": "gelpack 750 GR",
            "12000914": "gelpack 750 GR",
            "12000913": "gelpack 750 GR",
            "12000250": "gelpack 750 GR",
            "12000190": "gelpack 750 GR",
            "12001123": "gelpack 750 GR",
            "12001124": "gelpack 750 GR",
            "12001190": "gelpack 750 GR",
            "12001191": "gelpack 750 GR",
            "12001192": "gelpack 750 GR",
            "12001193": "gelpack 750 GR",
            "12000912": "gelpack 1000 KG",
            "12000458": "gelpack 1000 KG",
            "12000387": "gelpack 2000 KG",
            "12000368": "gelpack 2000 KG",
            "12000687": "gelpack 2000 KG"
        }

        # Redimensionar la ventana de SAP
        time.sleep(1)  # Pausa para dar tiempo a cargar la ventana
        try:
            session.findById("wnd[0]").resizeWorkingPane(118, 36, False)
            print("Ventana redimensionada correctamente.")
        except Exception as e:
            print(f"Error al redimensionar la ventana: {e}")
            return

        # Intentar obtener los códigos de material de todas las filas posibles
        fila = 0
        while True:
            try:
                material_code = session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,{fila}]").text
                print(f"Código de material obtenido en fila {fila}: {material_code}")

                # Determinar el tipo de gelpack según el material
                if material_code in material_gelpack:
                    material = material_gelpack[material_code]
                    print(f"Material seleccionado: {material}")
                else:
                    print(f"Error: No se encontró el tipo de gelpack para el material {material_code}")
                    break

                # Enfocar el campo de subposición antes de entrar a Datos Adicionales A
                time.sleep(1)
                session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,{fila}]").setFocus()
                session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,{fila}]").caretPosition = 8
                session.findById("wnd[0]").sendVKey(2)
                print("Campo de subposición enfocado correctamente.")

                # Navegar a la pestaña de Datos Adicionales A
                time.sleep(1)
                session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13").select()
                print("Navegación a Datos Adicionales A realizada correctamente.")

                # Asignar termógrafo interno y externo solo al primer producto (fila 1,0)
                if fila == 0:
                    time.sleep(1)  # Pausa para dar tiempo a cargar el elemento
                    try:
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR2").key = "001"  # Termógrafo interno
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR3").key = "001"  # Termógrafo externo
                        session.findById("wnd[0]").sendVKey(0)
                        print("Termógrafos asignados correctamente a la primera fila.")
                    except Exception as e:
                        print(f"Error al asignar los termógrafos: {e}")
                        return

                # Asignar cantidad de hielo según el material usando el código específico
                time.sleep(1)  # Pausa para dar tiempo a cargar el elemento
                if material == "gelpack 250 GR":
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").key = "008"
                elif material == "gelpack 500 GR":
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").key = "009"
                elif material == "gelpack 750 GR":
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").key = "010"
                elif material == "gelpack 1000 KG":
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").key = "011"
                elif material == "gelpack 2000 KG":
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").key = "015"
                else:
                    print(f"Error: No se encontró el código para el material {material}")
                    break
                session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR4").setFocus()
                session.findById("wnd[0]").sendVKey(0)
                print("Cantidad de hielo asignada correctamente.")

                # Enfocar y continuar
                time.sleep(1)  # Pausa para dar tiempo a cargar el elemento
                session.findById("wnd[0]").sendVKey(3)
                print("Enfoque y continuación realizados correctamente.")

                # Incrementar la fila para la siguiente iteración
                fila += 1

            except Exception as e:
                print(f"No se encontró más material en la fila {fila}. Error: {e}")
                break

    except Exception as e:
        print(f"Ocurrió un error: {e}")

if __name__ == "__main__":
    main()
