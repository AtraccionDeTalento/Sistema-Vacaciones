# -*- coding: utf-8 -*-
"""Mapa Area/Departamento -> Business Partner, derivado del organigrama de Talento y Cultura.
Granularidad: primero AREA (Nombre Area del maestro), luego DEPARTAMENTO (Nombre Departamento).
Devuelve uno de: 'Carlos Jara','Fatima Salazar','Lesley Reyes','Melissa Higa','REVISAR'."""
import re, unicodedata

def norm(s):
    if s is None: return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s

# --- Nivel AREA (gana sobre departamento). Claves normalizadas (sin acentos, mayus). ---
AREA_BP = {
    # VICERRECTORADO ACADEMICO -> facultades / direcciones
    "FACULTAD DE CIENCIAS DE LA SALUD": "Melissa Higa",
    "CENTER FOR LANGUAGE STUDIES": "Melissa Higa",          # CLS
    "CENTER FOR GLOBAL EDUCATION": "Melissa Higa",          # CGE
    "FACULTAD DE INGENIERIA E INTELIGENCIA ARTIFICIAL": "Fatima Salazar",
    "FACULTAD DE DERECHO": "Fatima Salazar",
    "FACULTAD DE EDUCACION": "Fatima Salazar",
    "DIRECCION DE ESTUDIOS GENERALES": "Fatima Salazar",
    "DIRECCION DE CALIDAD ACADEMICA, DOCENTE Y CURRICULAR": "Fatima Salazar",
    "DIRECCION DE CALIDAD ACADEMICA DOCENTE Y CURRICULAR": "Fatima Salazar",
    "FACULTAD DE CIENCIAS EMPRESARIALES": "Lesley Reyes",
    "FACULTAD DE ARTES Y HUMANIDADES": "Lesley Reyes",
    "FACULTAD ADM HOTELE TURIS Y GASTRO": "Lesley Reyes",
    "FACULTAD DE ARQUITECTURA URBANISMO Y TERRITORIO": "Lesley Reyes",
    "FACULTAD DE COMUNICACION": "Lesley Reyes",
    "DIRECCION PGEX": "Lesley Reyes",                        # Pregrado Ejecutivo
    # RECTORADO -> direcciones / vicepresidencias
    "VICEPRESIDENCIA DE INTERNACIONALIZACION": "Melissa Higa",
    "GERENCIA DE EMPLEABILIDAD E INNOVACION EDUCATIVA": "Carlos Jara",
    "SECRETARIA GENERAL": "Carlos Jara",
    "DIRECCION DE SOSTENIBILIDAD": "Lesley Reyes",
    # SEDES DESCENTRALIZADAS / VP SENIOR
    "INSTITUTO DE EMPRENDEDORES": "Carlos Jara",
    "GERENCIA DE EDUCACION VIRTUAL Y TEC. EDUCATIVA": "Melissa Higa",
    "GERENCIA DE EDUCACION VIRTUAL Y TEC EDUCATIVA": "Melissa Higa",
    # Auditoria (si aparece como area)
    "AUDITORIA": "Carlos Jara",
    "AUDITORIA INTERNA": "Carlos Jara",
}

# --- Nivel DEPARTAMENTO (fallback cuando el area no decide). ---
DEP_BP = {
    "ESCUELA DE POSGRADO": "Carlos Jara",
    "GERENCIA COMERCIAL": "Carlos Jara",
    "GERENCIA DE TALENTO Y CULTURA": "Carlos Jara",
    "VICEPRESIDENCIA LEGAL": "Carlos Jara",
    "VICERRECTORADO DE INVESTIGACION": "Fatima Salazar",
    "GERENCIA DE ADMINISTRACION Y SEGURIDAD": "Lesley Reyes",
    "VICEPRESIDENCIA DE FINANZAS": "Lesley Reyes",
    "GERENCIA DE MODO USIL": "Melissa Higa",
    "GERENCIA DE TECNOLOGIA DE INFORMACION": "Melissa Higa",
    "VICEPRESIDENCIA DE COMUNICACIONES E INNOVACION DIG": "Melissa Higa",
    "VICEPRESIDENCIA DE COMUNICACIONES E INNOVACION DIGITAL": "Melissa Higa",
    "VICEPRESIDENCIA DE ASUNTOS CORPORATIVOS": "Melissa Higa",
    "GERENCIA GENERAL": "Melissa Higa",
    "GERENCIA GENERAL ADJUNTA": "Melissa Higa",
    "VICEPRESIDENCIA SENIOR": "Melissa Higa",
    "DIRECCION GENERAL DE SEDES DESCENTRALIZADAS": "Carlos Jara",
    # NO en organigrama -> REVISAR
    "DIRECCION GENERAL DEL CSIR": "REVISAR",
    "PRESIDENCIA EJECUTIVA": "REVISAR",
    "GERENCIA DE PROYECTOS - SEDE PERU": "REVISAR",
}

# --- Reglas por SUBCADENA de AREA (robustas a nombres truncados en el maestro). ---
# Se evaluan tras AREA_BP exacto y antes de DEP_BP. Orden = prioridad.
AREA_CONTAINS = [
    ("CALIDAD ACADEMICA", "Fatima Salazar"),     # "...DOCENTE Y CURRICUL" truncado
    ("EMPLEABILIDAD", "Carlos Jara"),            # Empleabilidad y Alumni / Innovacion Educativa
    ("FONDO EDITORIAL", "Fatima Salazar"),       # Vicerrectorado de Investigacion
    ("EMPRENDIMIENTO", "Carlos Jara"),           # VP Emprendimiento (Rectorado)
    ("INSTITUTO DE EMPRENDEDORES", "Carlos Jara"),
    ("CIENCIAS DE LA SALUD", "Melissa Higa"),
    ("INTERNACIONALIZACION", "Melissa Higa"),
    ("EDUCACION VIRTUAL", "Melissa Higa"),
    ("SOSTENIBILIDAD", "Lesley Reyes"),
    ("SECRETARIA GENERAL", "Carlos Jara"),
    ("AUDITORIA", "Carlos Jara"),
]

def bp_de(departamento, area):
    a = norm(area); d = norm(departamento)
    if a in AREA_BP: return AREA_BP[a], f"area:{a}"
    for sub, bp in AREA_CONTAINS:
        if sub and sub in a: return bp, f"area~{sub}"
    if d in DEP_BP:  return DEP_BP[d], f"dep:{d}"
    return "REVISAR", f"sin_regla(dep:{d}|area:{a})"
