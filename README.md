# Tesis-Dapica

Materiales relativos a la [tesis doctoral de Rubén Dapica](https://narratech.com/es/project/tesis-dapica/), que deben complementarse también con [el repositorio de la ontología original EBTOnto](https://github.com/narratech/EBTOnto).

## ASRS
Base de datos de incidentes aéreos en formato CSV. Al ser algo pesada está dentro del almacén de Git LFS asociado a este repositorio.

## Generator
Código App Script y hojas de cálculo utilizadas para generar la base de conocimiento en función de la entrada tomada del ASRS.

La nueva versión de la ontología EBTOnto está también subida aquí, ahora se llama EBTOnto.ofn porque ha cambiado de formato.

Se incluye también catalog-v001.xml, que debe estar en el mismo directorio que la ontología, porque Protègè tiene problemas importando ontologías en formato OFN (OWL Funcional).  

Para usar el generador de individuos (de la base de conocimiento) hay que juntar toda la info que nos interesa del ASRS en una única hoja de cálculo, si puede ser... la 'Entrada' y pones el identificador de dicha hoja de cálculo en la primera línea del programa de App Script (variable idEntrada), para que lea la entrada de allí. 
Se deberían generar N ficheros resultantes, que se tendrán que meter juntos en un fichero de texto (ASRSkg.ofn, se puede llamar… aunque realmente no importa) y ese fichero es el que tratas de abrir con Protègè. Debería cargar todo sin fallos e incluso permitirte activar el razonador sin problemas de inconsistencias :-)  

Finalmente, una vez esté todo cargado, validado y clasificado en Protègè, vendría la parte de hacer consultas con SPARQL, por ejemplo. 

## Jupyter Notebook
Código en Python para realizar la parte de aprendizaje automático.
