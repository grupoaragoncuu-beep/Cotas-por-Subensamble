import pythoncom
import win32com.client
from inventor_com import conectar_inventor

def main():
    pythoncom.CoInitialize()
    try:
        inv_app = conectar_inventor()
        if not inv_app:
            print("No se pudo conectar a Inventor.")
            return
            
        print("Conectado a Inventor.")
        print(f"Total documentos abiertos: {inv_app.Documents.Count}")
        
        for i in range(1, inv_app.Documents.Count + 1):
            doc = inv_app.Documents.Item(i)
            print(f"[{i}] {doc.DisplayName} - Tipo: {doc.DocumentType}")
            
            # If it's an assembly, let's look at its leaves
            if doc.DocumentType == 12291 or "Assembly" in doc.DisplayName:
                try:
                    ens = win32com.client.CastTo(doc, "AssemblyDocument")
                    if ens:
                        occs = ens.ComponentDefinition.Occurrences.AllLeafOccurrences
                        print(f"  -> AssemblyDocument cast exitoso. Hojas: {occs.Count}")
                        for j in range(1, min(occs.Count + 1, 10)): # print first 9
                            occ = occs.Item(j)
                            print(f"    - Occ: {occ.Name} | Tipo: {occ.DefinitionDocumentType} | Suppressed: {occ.Suppressed}")
                except Exception as e:
                    print(f"  -> Error al inspeccionar ensamble: {e}")
                    
    except Exception as e:
        print(f"EXCEPTION: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
