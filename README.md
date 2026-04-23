# Smallworld GIS Automation Platform

**Plataforma empresarial para automatización de procesos geoespaciales con integración AWS Cloud**

[![AWS](https://img.shields.io/badge/AWS-Certified-FF9900?style=for-the-badge&logo=amazon-aws)](https://aws.amazon.com)
[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python)](https://python.org)
[![Terraform](https://img.shields.io/badge/Terraform-1.0+-7B42BC?style=for-the-badge&logo=terraform)](https://terraform.io)
[![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)](LICENSE)

---

##  Tabla de Contenidos

- [Descripción General](#descripción-general)
- [Arquitectura](#arquitectura)
- [Estructura del Proyecto](#estructura-del-proyecto)
- [Tecnologías Utilizadas](#tecnologías-utilizadas)
- [Licencia](#licencia)

---

## Descripción General

**Smallworld GIS Automation Platform** es una solución empresarial diseñada para automatizar y orquestar procesos geoespaciales. La plataforma integra **Smallworld GIS** con servicios **AWS Cloud**, proporcionando:

- **Sincronización bidireccional** entre Smallworld y bases de datos PostgreSQL/PostGIS
- **Pipelines ETL automatizados** para procesamiento de datos geoespaciales
- **Validación topológica** y control de calidad de datos
- **Monitoreo en tiempo real** con dashboards CloudWatch
- **Infrastructure as Code** para despliegues consistentes y reproducibles

### Casos de Uso Principales

| Caso de Uso | Descripción |
|-------------|-------------|
| **Gestión Masiva** | Desarrollo scripts por control orientado a objetos 
| **Sincronización GIS-DB** | Replicación de cambios entre Smallworld y base de datos corporativa |
| **Control de Calidad** | Validación automática de integridad topológica y atributiva |
| **Reportes Automáticos** | Generación de reportes de calidad y auditoría geoespacial |
| **Migración de Datos** | ETL para ingesta de datos desde múltiples fuentes (SHP, DWG, Excel) |

---


 RDS PostgreSQL + PostGIS - Datos geoespaciales (puntos, líneas, polígonos)  - Auditoría y tracking de cambios 


 - Smallworld GIS                  
 - Magik/Ruby scripts            






## Arquitectura - SMALLWORLD GIS PLATFORM (Generado: 2026-04-22)


GS
├── src/                                    
│   ├── core/                               
│   │   ├── database/                       
│   │   │   └── migrations/                 
│   │   ├── geospatial/                     
│   │   │   ├── processors/  # Procesadores GIS
│   │   │   └── validators/  # Validadores topológicos
│   │   └── sync/                           
│   │       ├── engines/                    
│   │       └── handlers/                   
│   │
│   ├── lambda/                            
│   │   ├── handlers/                       
│   │   │   ├── quality/                    
│   │   │   └── sync/                       
│   │   └── layers/                         
│   │       └── shared/                     
│   │
│   ├── pipelines/                          
│   │   ├── etl/                            
│   │   │   ├── extract/                   
│   │   │   ├── transform/                 
│   │   │   └── load/                       
│   │   └── quality/                        
│   │       ├── reporters/                  
│   │       └── validators/                 
│   │
│   ├── utils/                             
│   │   ├── helpers/                        
│   │   ├── python/                         
│   │   └── validators/                     
│   │
│   └── private_collections/                
│       ├── AttrATC/                        
│       └── InfraUndergroundRoute/          
│
├── infrastructure/                        
│   ├── terraform/                          
│   │   ├── environments/                  
│   │   │   ├── dev/                       
│   │   │   │   └── main.tf                 
│   │   │   ├── staging/                    
│   │   │   └── prod/                      
│   │   └── modules/                        
│   │       ├── database/                
│   │       ├── monitoring/               
│   │       └── smallworld/                 
│   │
│   ├── cloudformation/                    
│   │   ├── nested/                       
│   │   └── stacks/                      
│   │
│   ├── aws/                               
│   │   ├── cloudformation/                
│   │   └── lambda/                         
│   │
│   └── docker/                           
│
├── config/                                
│   ├── environments/                      
│   │   ├── development/                    
│   │   │   └── smallworld.json             
│   │   ├── staging/                       
│   │   │   └── smallworld.json
│   │   └── production/                    
│   │       └── smallworld.json
│   │
│   ├── security/                         
│   │   ├── credentials/                   
│   │   ├── encryption/                    
│   │   └── secrets/                       
│   │
│   └── config.json                       
│
├── docs/                                   
│   ├── api/                                
│   │   ├── rest/                        
│   │   └── websocket/                      
│   │
│   ├── architecture/                     
│   │   ├── decisions/                                        
│   │   └── diagrams/    
│   │
│   ├── operations/    # Documentación operativa
│   │   ├── incidents/                     
│   │   └── runbooks/  # Procedimientos operativos
│   │
│   ├── security/                         
│
├── database/                            
│   ├── migrations/                     
│   │   ├── 001_consultas.sql             
│   │   ├── 001_depurar_asphia.sql       
│   │   └── 001_pruebas.sql                 
│   │
│   ├── output/                            
│   │   ├── 01-Bulk export of buildings.xlsx
│   │   ├── 01-Bulk export of fibers.xlsx
│   │   ├── 01-Bulk export of ports.xlsx
│   │   ├── 01-Bulk export of splices.xlsx
│   │   ├── 01-Bulk Merge_Centrales.xlsx
│   │   ├── 02-Individual export of buildings.xlsx
│   │   ├── 02-Individual export of fibers.xlsx
│   │   ├── 02-Individual export of ports.xlsx
│   │   ├── 02-Individual export of splices.xlsx
│   │   ├── 03-2025-08-31_BASE DE ODF´S.xlsx
│   │   ├── 03-ASPHIA_CLIENTES.xlsx
│   │   ├── 03-LINEA CABLES PRIORIZADOS.xlsx
│   │   ├── 03-TRAZE_MS_CONNECTION.xlsx
│   │   ├── 04-AuditoriaCalidadVisualAtributiva.xlsm
│   │   ├── 04-FIBRAS-ELIMINADAS.xlsx
│   │   ├── 04-Sync_SW_SP.xlsx
│   │   ├── MG-01-Bulk export of ports.xlsx
│   │   ├── MG-01-Bulk Merge_Centrales.xlsx
│   │   └── PROCESO-04-AuditoriaCalidadVisualAtributiva.xlsm
│   │
│   ├── reference/                         
│   └── smallworld/                         
│       └── colleccion/
│           └── find_output/            
│
├── tests/                              
│   ├── e2e/                             
│   ├── integration/                       
│   │   └── pipelines/                     
│   ├── unit/                         
│   │   └── core/                          
│   └── 01_TR_INDEX                        
│
├── scripts/                               
│   ├── database/                      
│   │   ├── migrations/             
│   │   └── seeds/                   
│   ├── deployment/                      
│   ├── monitoring/                       
│   │   ├── alerts/                         
│   │   ├── cloudwatch/                 
│   │   └── dashboards/                    
│   ├── setup/                    
│   ├── shortcuts/                        
│   └── sw_trz/                         
│
├── .github/                               
│   ├── workflows/                     
│   ├── ISSUE_TEMPLATE/            
│   └── PULL_REQUEST_TEMPLATE/          
│
├── README.md                            
├── requirements.txt                     
├── requirements-dev.txt            
├── setup.py                                
├── Makefile                              
└── .gitignore                    