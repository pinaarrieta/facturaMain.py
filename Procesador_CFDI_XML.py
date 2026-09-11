#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de CFDI 4.0 por lotes.
Toma todos los archivos XML de una carpeta y genera un único archivo CSV.
"""

import xml.etree.ElementTree as ET
import csv
from typing import Dict, Any, List
import os
import glob # Librería para buscar archivos con patrones
import sys
import platform

# Namespaces comunes en CFDI 4.0 + TimbreFiscalDigital de la SHCP-SAT en México
NAMESPACES = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd':  'http://www.sat.gob.mx/TimbreFiscalDigital',
    # Agrega otros namespaces si son necesarios (e.g., Pago, Nómina), este programa solo considera ingresos
}

# --- 1. Función de Parseo de XML, desmenuza cada elemento del árbol XML---

def parse_cfdi_xml(file_path: str) -> Dict[str, Any]:
    """
    Extrae todos los datos relevantes de un CFDI 4.0 (XML) y los devuelve como dict.
    Simplificada para solo tomar la ruta del archivo.
    """
    try:
        # Parseo
        tree = ET.parse(file_path)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"ERROR de Parseo en {file_path}: {e}")
        return {}

    def atts(elem):
        """Función auxiliar para obtener atributos de un elemento."""
        return dict(elem.attrib) if elem is not None else {}

    result: Dict[str, Any] = {}
    result['Comprobante'] = atts(root)

    # Emisor
    emisor = root.find('cfdi:Emisor', NAMESPACES)
    result['Emisor'] = atts(emisor)

    # Receptor
    receptor = root.find('cfdi:Receptor', NAMESPACES)
    result['Receptor'] = atts(receptor)

    # Conceptos (lista)
    conceptos_list: List[Dict[str, Any]] = []
    conceptos = root.find('cfdi:Conceptos', NAMESPACES)
    if conceptos is not None:
        for c in conceptos.findall('cfdi:Concepto', NAMESPACES):
            c_dict = atts(c)
            
            # Extracción de Impuestos a nivel Concepto (Traslados y Retenciones)
            c_dict['Traslados'] = []
            c_dict['Retenciones'] = []
            
            impuestos = c.find('cfdi:Impuestos', NAMESPACES)
            if impuestos is not None:
                traslados_node = impuestos.find('cfdi:Traslados', NAMESPACES)
                if traslados_node is not None:
                    c_dict['Traslados'] = [atts(t) for t in traslados_node.findall('cfdi:Traslado', NAMESPACES)]

                retenciones_node = impuestos.find('cfdi:Retenciones', NAMESPACES)
                if retenciones_node is not None:
                    c_dict['Retenciones'] = [atts(r) for r in retenciones_node.findall('cfdi:Retencion', NAMESPACES)]
            
            conceptos_list.append(c_dict)
            
    result['Conceptos'] = conceptos_list

    # Impuestos a nivel comprobante
    impuestos_tot = root.find('cfdi:Impuestos', NAMESPACES)
    result['Impuestos'] = atts(impuestos_tot)
    
    #Extraer TotalImpuestosTrasladados
    result['TotalImpuestosTrasladados'] = ''
    if impuestos_tot is not None:
        result['TotalImpuestosTrasladados'] = impuestos_tot.attrib.get('TotalImpuestosTrasladados', '')

    # Complementos: TimbreFiscalDigital (tfd)
    complementos = root.find('cfdi:Complemento', NAMESPACES)
    result['Complemento'] = {}
    if complementos is not None:
        tfd = complementos.find('tfd:TimbreFiscalDigital', NAMESPACES)
        if tfd is not None:
            result['Complemento']['TimbreFiscalDigital'] = atts(tfd)

    # Resumen para fácil acceso
    uuid = result.get('Complemento', {}).get('TimbreFiscalDigital', {}).get('UUID')
    
    # Agregamos la ruta del archivo para referencia en el CSV final
    result['CFDI_RutaArchivo'] = file_path
    result['CFDI_UUID'] = uuid
    
    return result

# --- 2. Función para Aplanar los Datos de un XML a Filas de CSV ---

def normalize_data_for_csv(cfdi_data: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Toma los datos estructurados de un CFDI y los convierte en una lista
    plana de diccionarios, duplicando los datos del encabezado por cada concepto.
    """
    rows: List[Dict[str, str]] = []

    # 1. Preparar datos de Cabecera (Se usarán en cada fila)
    comprobante = cfdi_data.get('Comprobante', {})
    emisor = cfdi_data.get('Emisor', {})
    receptor = cfdi_data.get('Receptor', {})
    
    # Intenta obtener el UUID del Timbre (lo más importante)
    uuid = cfdi_data.get('CFDI_UUID', '')
    
    header_data = {
        # Datos del Archivo
        'CFDI_UUID': uuid,
        'CFDI_NombreArchivo': os.path.basename(cfdi_data.get('CFDI_RutaArchivo', '')),
        
        # Datos del Comprobante
        'CFDI_Version': comprobante.get('Version', ''),
        'CFDI_Serie': comprobante.get('Serie', ''),
        'CFDI_Folio': comprobante.get('Folio', ''),
        'CFDI_Fecha': comprobante.get('Fecha', ''),
        'CFDI_SubTotal': comprobante.get('SubTotal', ''),
        'CFD_Descuento': comprobante.get('Descuento', ''), #02/03/2026
        'CFDI_TotalImpuestosTrasladados': cfdi_data.get('TotalImpuestosTrasladados', ''),
        'CFDI_Total': comprobante.get('Total', ''),
        'CFDI_Moneda': comprobante.get('Moneda', ''),
        'CFDI_TipoComprobante': comprobante.get('TipoDeComprobante', ''),
        'CFDI_MetodoPago': comprobante.get('MetodoPago', ''),
        'CFDI_FormaPago': comprobante.get('FormaPago', ''),

        # Datos del Emisor
        'EMISOR_Rfc': emisor.get('Rfc', ''),
        'EMISOR_Nombre': emisor.get('Nombre', ''),
        'EMISOR_RegimenFiscal': emisor.get('RegimenFiscal', ''),

        # Datos del Receptor
        'RECEPTOR_Rfc': receptor.get('Rfc', ''),
        'RECEPTOR_Nombre': receptor.get('Nombre', ''),
        'RECEPTOR_UsoCFDI': receptor.get('UsoCFDI', ''),
    }

    conceptos = cfdi_data.get('Conceptos', [])

    if not conceptos:
        # Caso sin conceptos (ej. CFDI de Pago o Nómina):
        # Creamos una fila de sólo encabezado para registrar el CFDI.
        rows.append({**header_data, 'CONCEPTO_Descripcion': '[Sin Conceptos / CFDI de Pago]'})
        return rows
    
    # 2. Iterar sobre cada Concepto y aplanarlo
    for concepto in conceptos:
        row = header_data.copy()
        
        # Datos del Concepto
        row.update({
            'CONCEPTO_NoIdentificacion': concepto.get('NoIdentificacion', ''),
            'CONCEPTO_Descripcion': concepto.get('Descripcion', ''),
            'CONCEPTO_ClaveProdServ': concepto.get('ClaveProdServ', ''),
            'CONCEPTO_Cantidad': concepto.get('Cantidad', ''),
            'CONCEPTO_ClaveUnidad': concepto.get('ClaveUnidad', ''),
            'CONCEPTO_Unidad': concepto.get('Unidad', ''),
            'CONCEPTO_ValorUnitario': concepto.get('ValorUnitario', ''),
            'CONCEPTO_Importe': concepto.get('Importe', ''),
            'CONCEPTO_Descuento': concepto.get('Descuento', ''),
        })
        
        # --- Extracción de Impuestos Trasladados (IVA y IEPS) ---
        iva_traslado_base = ''
        iva_traslado_tasa = ''
        iva_traslado_importe = ''
        
        ieps_traslado_base = ''
        ieps_traslado_tasa = ''
        ieps_traslado_importe = ''
        
        # Iteramos sobre todos los traslados del concepto para buscar IVA (002) e IEPS (003)
        for t in concepto.get('Traslados', []):
            impuesto_clave = t.get('Impuesto')
            
            if impuesto_clave == '002': # 002 = IVA
                 # Asignar los valores de IVA
                 iva_traslado_base = t.get('Base', '')
                 iva_traslado_tasa = t.get('TasaOCuota', '')
                 iva_traslado_importe = t.get('Importe', '')
                 
            elif impuesto_clave == '003': # 003 = IEPS
                 # Asignar los valores de IEPS
                 ieps_traslado_base = t.get('Base', '')
                 ieps_traslado_tasa = t.get('TasaOCuota', '')
                 ieps_traslado_importe = t.get('Importe', '')

        # Actualizar la fila con los datos de impuestos encontrados
        row.update({
            'IMPUESTO_IVA_Base': iva_traslado_base,
            'IMPUESTO_IVA_TasaOCuota': iva_traslado_tasa,
            'IMPUESTO_IVA_Importe': iva_traslado_importe,
            'IMPUESTO_IEPS_Base': ieps_traslado_base,      # AÑADIDO: Base de IEPS
            'IMPUESTO_IEPS_TasaOCuota': ieps_traslado_tasa, # AÑADIDO: Tasa/Cuota de IEPS
            'IMPUESTO_IEPS_Importe': ieps_traslado_importe, # AÑADIDO: Importe de IEPS
        })
        
        rows.append(row)
        
    return rows

# --- 3. Función Principal de Lote ---

def batch_extractor():
    """
    Solicita la carpeta de entrada y el archivo CSV de salida, procesa
    todos los XML y escribe el resultado consolidado.
    """
    limpiar_pantalla()
    print('\n**********************************************************************************')
    print('********************* Extractor de Lotes CFDI 4.0 a CSV Único **********************')
    print('**********************************************************************************\n')
    
    # 1. Solicitar ruta de la carpeta de XMLs
    folder_path = input("Ingrese la RUTA COMPLETA de la carpeta con sus archivos .xml: ").strip()
    
    if not os.path.isdir(folder_path):
        print(f"\n❌ Error: La ruta '{folder_path}' no es una carpeta válida.")
        input("Presione Enter para salir...")
        return
        
    # 2. Buscar archivos XML
    # Se utiliza glob para encontrar todos los archivos con extensión .xml o .XML, en windows es indistinto, 
    # pero en Linux si importan las mayusculas y minusculas
    xml_files = glob.glob(os.path.join(folder_path, '*.xml')) #+ glob.glob(os.path.join(folder_path, '*.XML'))
    
    if not xml_files:
        print(f"\n⚠️ Advertencia: No se encontraron archivos XML en '{folder_path}'.")
        input("Presione Enter para salir...")
        return
        
    print(f"\n✅ Se encontraron {len(xml_files)} archivos XML listos para procesar.")

    # 3. Solicitar nombre del archivo de salida
    csv_out_name = input("Ingrese el NOMBRE para el archivo CSV de salida (ej. Reporte_CFDI.csv): ").strip()
    if not csv_out_name.lower().endswith('.csv'):
        csv_out_name += '.csv'
        
    csv_output_path = os.path.join(folder_path, csv_out_name)
    
    all_csv_rows: List[Dict[str, str]] = []
    
    # 4. Procesamiento de Lote
    print("\nIniciando procesamiento...")
    
    for i, xml_path in enumerate(xml_files):
        print(f"[{i+1}/{len(xml_files)}] Procesando: {os.path.basename(xml_path)}...")
        
        # Parsear XML
        cfdi_data = parse_cfdi_xml(xml_path)
        
        if cfdi_data:
            # Aplanar y agregar al listado general
            flattened_rows = normalize_data_for_csv(cfdi_data)
            all_csv_rows.extend(flattened_rows)

    # 5. Escritura del CSV Consolidado
    if not all_csv_rows:
        print("\n⚠️ No se pudieron extraer datos de ningún archivo. El CSV no será creado.")
        return
        
    # Obtener todos los nombres de campo (columnas) a partir de la primera fila,
    # que debe contener todos los campos definidos.
    fieldnames = list(all_csv_rows[0].keys())
    
    try:
        with open(csv_output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(all_csv_rows)
            
        print(f"\n==================================================================================")
        print(f"🎉 ¡Procesamiento Completo!")
        print(f"Archivo CSV consolidado creado con éxito en:\n-> {csv_output_path}")
        print(f"Total de filas (conceptos/facturas): {len(all_csv_rows)}")
        print(f"==================================================================================")

    except Exception as e:
        print(f"\n❌ Error al escribir el archivo CSV: {e}")
        
    input("\nPresione Enter para salir...")


# --- Funciones Auxiliares (mantienen su utilidad) ---

def limpiar_pantalla():
    """Limpia la consola para una mejor visualización."""
    if platform.system() == "Windows":
        os.system("cls")
    else: # Linux y macOS
        os.system("clear")

# --- Programa Principal ---
if __name__ == "__main__":
    batch_extractor()
