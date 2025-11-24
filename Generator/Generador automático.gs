/*********************************************************************************************************
 * ASRS Knowledge Graph Automatic Generator for the EBTOntology 
 * Versión 1.2
 * 
 * Generador automático de individuos para EBTOnto (Base de conocimiento en lenguaje OWL DL y formato OWL Functional) 
 * partiendo desde la base de datos ASRS (ficheros Excel descargados y unificados en una hoja de cálculo Google Sheets)
 *
 * Tesis doctoral de Rubén Dapica
 * Software de apoyo creado por Federico Peinado
 ********************************************************************************************************/
 
/*********************************************************************************************************
 * Función principal para generar los individuos para la ontología a partir de las filas de la base de datos. Posibles mejoras:
 * - Mejorar la detección de palabras clave en columnas como Narrative, de texto libre, usando expresiones regulares u otras técnicas de análisis/procesamiento de lenguaje.
 * - Profundizar en la generación del grafo, creando subindividuos que tengan a su vez subpropiedades hacia subsubindividuos, etc. (Ej. distinguir entre Scenario y ASReport)
 * - Hacer que subindividuos aparezcan referidos como blanco de VARIAS propiedades de un individuo y cosas así (por poder, se podría modelar a los pilotos involucrados y todo eso lo del ASRS).
 * - Publicarla como aplicación web con la que un usuario puedas especificar el nombre(id) del fichero de entrada, del de traducción y del de salida a utilizar.
 * - Enrriquecer el comentario de cada individuo para que sea un resumen de lo que es y las propiedades que tiene, del informe ASRS del que viene, etc.
 * - Si se quejase GoogleDocs por hacer demasiados cambios seguidos, se puede ir grabando el resultado entre medias: programa.saveAndClose();  
 * - Si necesitas añadir contenido a pelo al informe se puede hacer por ejemplo así: cuerpo.appendListItem('Encabezado | ' + actividad[columna]).setGlyphType(DocumentApp.GlyphType.BULLET);
 * - Para depurar es buena idea usar Logger.log() y así imprimes en el registro de ejecución
 ********************************************************************************************************/
function generaIndividuos() {
  

  /***************************************************************************/
  // El ID de la hoja de cálculo (Google Sheets) que hace las veces de entrada
  // - Lo hemos tocado, la entrada habitual es esa que se llama 'Entrada', con este id: 1XbCX_tvOuhuB7VIu691RBh8qF_cXYjH3SSYxktB51R4
  // - Este otro id, por ejemplo, es de una hoja de cálculo que descargó Dapica: 1ioOoq9XshpDbyPu8WBzTnLZTjf1wYRacFAC58KnuvQU
  var idEntrada = '1IgDDN7MdqmG-Ya10l1V2vQJlR1l-pRYlsfWU7OXoKIY'; 
  /***************************************************************************/  


  // La lista con todos los IDs de las columnas de la tabla
  var listaIdColumnas = Sheets.Spreadsheets.Values.get(idEntrada, 'A2:CR2').values[0]; // Los IDs de los encabezados estarán justo en la segunda fila y son 96

  // La lista con todas las filas de contenido de la tabla de entrada
  // - Debería bastar con llegar a columna CR
  // - Se debería consultar previamente la longitud máxima ocupada en la hoja de cálculo. Por ahora está puesto este valor de 5000 filas, si se necesitan más se puede cambiar.
  var matrizIncidentes = Sheets.Spreadsheets.Values.get(idEntrada, 'A4:CU5004').values; 
  
  // El ID de la hoja de cálculo (Google Sheets) con las traducciones
  // - Normalmente es el fichero que se llama 'Traducciones', con este id: 1e-aAYHdDk7S36wmh-9lDga_snIKzYxeLH1hsVaLB7_0
  var idTraduccion = '1e-aAYHdDk7S36wmh-9lDga_snIKzYxeLH1hsVaLB7_0';

  // La lista con todas las filas de contenido de la tabla de traducciones
  // - Se debería consultar previamente la longitud máxima ocupada en la hoja de cálculo. Por ahora está puesto este valor de 1000 filas, si se necesitan más se puede cambiar. 
  var matrizTraducciones = Sheets.Spreadsheets.Values.get(idTraduccion, 'A2:V1002').values; 
  
  // El documento (Google Docs) que sirve de base para generar cada informe
  // - Normalmente es el fichero que se llama 'Plantilla para la salida', con este id: 1Z4R56SCYbPs8GL8wV5ingMXL2U8J19X7AtpYnPABTYw
  var idPlantillaInforme = '1Z4R56SCYbPs8GL8wV5ingMXL2U8J19X7AtpYnPABTYw'; 
  

  // ESTRUCTURA OWL DE LOS INDIVIDUOS  
  // Se usan las llaves {} para marcar los para los fragmentos sustituibles, ¡es importante que no aparezcan estos símbolos en los individuos de la ontología!  
  
  // Las IRIs (completas y abreviadas/cortas) que usamos para todos los tipos y propiedades tanto de la ontología como del grafo de conocimiento 
  // - Se podrían cambiar para que encaje con la web exacta donde la tengamos.
  // - Convendría usar abreviaturas para evitar que pese demasiado el texto generado.
  // - Ojo: Conviene cambiar todas las IRIs para que sean la nuestra, no de Web Protègè y también usar nombres legibles de tipos y propiedades, en lugar de IDs automáticos; 
  //        así como gestionar la importación desde Protègè de ficheros en local.    
  var iriOntoCorta = 'ebtonto'; // Ojo, no sé si hay que llamarla como antes, EBTOntology
  var iriOntoPrefix = 'ebt';
  var iriOnto = 'https://narratech.com/document' + '/' + iriOntoCorta;
  var iriCorta = 'ASRSkg'; // Abreviatura de ASRSKnowledgeGraph
  var iri = iriOnto + '/' + iriCorta; // Estoy colocando el grafo de conocimiento un nivel por debajo de la ontología... pero bueno, es como me lo imagino publicado en la web de Narratech


  /***************************************************************************/
  // Las plantillas RDF/XML con fragmentos sustituibles
  /***************************************************************************/

/*
  // Encabezado general del fichero
  var plantillaEncabezado = 
      '<?xml version="1.0"?>' + '\n' +
      '<rdf:RDF xmlns="' + iri + '#"' + '\n' +
      '    xml:base="' + iri + '"' + '\n' +
      '    xmlns:owl="http://www.w3.org/2002/07/owl#"' + '\n' +
      '    xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"' + '\n' +
      '    xmlns:xml="http://www.w3.org/XML/1998/namespace"' + '\n' +
      '    xmlns:xsd="http://www.w3.org/2001/XMLSchema#"' + '\n' +
      '    xmlns:rdfs="http://www.w3.org/2000/01/rdf-schema#"' + '\n' +
      '    xmlns:skos="http://www.w3.org/2004/02/skos/core#"' + '\n' +
            '    xmlns:skos="http://www.w3.org/2004/02/skos/core#"' + '\n' +
      '    xmlns:' + iriOntoPrefix + '="' + iriOnto + '#">' + '\n' +     
      '    <owl:Ontology rdf:about="' + iri + '">' + '\n' +
      '        <owl:imports rdf:resource="' + iriOnto + '"/>' + '\n' +
      '    </owl:Ontology>' + '\n\n';

  // Encabezado de cada individuo
  var plantillaEncabezadoIndividuos = 
      '    <!-- ' + '\n' +
      '    ///////////////////////////////////////////////////////////////////////////////////////' + '\n' +
      '    //' + '\n' +
      '    // Individuals' + '\n' +
      '    //' + '\n' +
      '    ///////////////////////////////////////////////////////////////////////////////////////' + '\n' +
      '    -->' + '\n\n';

  // Contenido de cada individuo
  // - No meter los comentarios y las etiquetas para reducir la cantidad de texto     
  // - ATRIVUTO se escribe con V porque \B parece ser un carácter especial que da problemas a la hora de reconocerlo en Google App Script 
  var plantillaIndividuo =
  //  '    <!-- ' + iri + '#' + '{ID}_{NUM} -->' + '\n' +
      '    <owl:NamedIndividual rdf:about="' + iri + '#' + '{ID}_{NUM}">' + '\n' +
      '{TIPOS}' +
      '{ATRIVUTOS}' +
      '{PROPIEDADES}' +
  //  '        <rdfs:label xml:lang="en">{ID}_{NUM}</rdfs:label>' + '\n' +
  //  '        <rdfs:label xml:lang="es">{ID}_{NUM}</rdfs:label>' + '\n' +
      '    </owl:NamedIndividual>' + '\n\n';

  // Contenido de cada subindividuo    
  // - Básicamente es como un individuo, pero no se les meten atributos ni propiedades, por ahora
  var plantillaSubindividuo = 
      '    <owl:NamedIndividual rdf:about="' + iri + '#' + '{IDAUX}_{NUMAUX}">' + '\n' +
      '{TIPOS}' +  
      '    </owl:NamedIndividual>' + '\n\n';

  // Tipo del individuo
  var plantillaTipo =
      '        <rdf:type rdf:resource="' + iriOnto + '#' + '{TIPO}"/>' + '\n';

  // Cada atributo del individuo
  var plantillaAtributo =
      '        <' + iriOntoCorta + ':{ATRIVUTO} rdf:datatype="{TIPO_ATRIVUTO}">' + '\n' +
      '{VALOR_ATRIVUTO}' + '\n' +
      '</' + iriOntoCorta + ':{ATRIVUTO}>' + '\n';

  // Cada propiedad del individuo
  var plantillaPropiedad =
      '        <' + iriOntoCorta + ':{PROPIEDAD} rdf:resource="' + iri + '#' + '{ID}_{NUM}"/>' + '\n'; 

  // Finalización general del fichero
  var plantillaFinal =
      '</rdf:RDF>' + '\n' +
      '<!-- Generated by the ASRS Knowledge Graph Automatic Generator for the EBTOntology (version 1.2) -->'; // Podría añadir la fecha y hora de generación, y leer la versión de una constante
*/

  /***************************************************************************/
  // Las plantillas OWL Functional con fragmentos sustituibles
  /***************************************************************************/

  // Encabezado general del fichero
  var plantillaEncabezado = 
      'Prefix(:=<' + iri + '#>)' + '\n' +  
      'Prefix(' + iriOntoPrefix + ':=<' + iriOnto + '#>)' + '\n' + 
      // Estos prefix no son del todo necesarios... pero a la mínima que quieras hacer algo con RDF o XSD Protègè te los va a meter automáticamente
      'Prefix(owl:=<http://www.w3.org/2002/07/owl#>)' + '\n' + 
      'Prefix(rdf:=<http://www.w3.org/1999/02/22-rdf-syntax-ns#>)' + '\n' + 
      'Prefix(xml:=<http://www.w3.org/XML/1998/namespace>)' + '\n' + 
      'Prefix(xsd:=<http://www.w3.org/2001/XMLSchema#>)' + '\n' + 
      'Prefix(rdfs:=<http://www.w3.org/2000/01/rdf-schema#>)' + '\n' + 
      '\n' +     
      'Ontology(<' + iri + '/>' + '\n' + 
      'Import(<' + iriOnto + '/>)' + '\n' + 
      '\n';    

  // Encabezado de cada individuo (Protègè suele colocar las Declaration incluso antes de este comentario)
  var plantillaEncabezadoIndividuos = 
      '############################' + '\n' +
      '#   Named Individuals' + '\n' +
      '############################' + '\n' +
      '\n';

  // Contenido de cada individuo
  // - No meter los comentarios y las etiquetas para reducir la cantidad de texto    
  // - ATRIVUTO se escribe con V porque \B parece ser un carácter especial que da problemas a la hora de reconocerlo en Google App Script 
  // - Ojo, realmente TIPOS es un único tipo... y estoy asumiendo que el orden en que pongamos estas partes dará igual
  var plantillaIndividuo =
  //  '# Individual: :{ID}_{NUM} (:{ID}_{NUM})' + '\n' +
      'Declaration(NamedIndividual(:{ID}_{NUM}))' + '\n' +
      '{TIPOS}' +
      '{ATRIVUTOS}' +
      '{PROPIEDADES}' +
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}")' + '\n' +
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}"@en)' + '\n' +
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}"@es)' + '\n' +
      '\n';

  // Contenido de cada subindividuo    
  // - No meter los comentarios y las etiquetas para reducir la cantidad de texto 
  // - Básicamente es como un individuo, pero no se les meten atributos ni propiedades, por ahora
  var plantillaSubindividuo = 
  //  '# Individual: :{ID}_{NUM} (:{ID}_{NUM})' + '\n' +
      'Declaration(NamedIndividual(:{ID}_{NUM}))' + '\n' +
      '{TIPOS}' + 
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}")' + '\n' +
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}"@en)' + '\n' +
  //  'AnnotationAssertion(rdfs:label :{ID}_{NUM} "{ID}_{NUM}"@es)' + '\n' +
      '\n';

  // Tipo del individuo
  var plantillaTipo =
     'ClassAssertion(' + iriOntoPrefix + ':{TIPO} :{ID}_{NUM})' + '\n';

  // Cada atributo del individuo
  // - Ojo, estoy añadiendo comillas para todos los atributos y realmente eso sólo debería hacerlo cuando es un atributo de tipo string
  var plantillaAtributo =
      'DataPropertyAssertion(' + iriOntoPrefix + ':{ATRIVUTO} :{ID}_{NUM} "' + '{VALOR_ATRIVUTO}"^^{TIPO_ATRIVUTO})' + '\n';

  // Cada propiedad del individuo ... en este caso necesitamos nombrar tanto al individuo como al subindividuo
  var plantillaPropiedad =
      'ObjectPropertyAssertion(' + iriOntoPrefix + ':{PROPIEDAD} :{ID}_{NUM} :{IDAUX}_{NUMAUX})' + '\n'; 

  // Finalización general del fichero
  var plantillaFinal =
      ')' + '\n' +
      '# Generated by the ASRS Knowledge Graph Automatic Generator for the EBTOntology (version 1.2)'; // Podría añadir la fecha y hora de generación, y leer la versión de una constante
 


//////////////////////////////////////////////////////////////////////////////// 


  
  
  // Se generará un informe por cada ejecución, con una parte con estadísticas de todo lo visto en la muestra de ASRS (cuantos incidentes tienen tales condiciones, en cuantos de ellos salen ciertas keywords que nosotros asociamos a competencias, etc.)
  // para ver la correlación; y otra parte donde se imprimirá la conversión a ABox de OWL, lo que se recomienda copiar y pegar dentro del fichero OWL de Protégé (o bien crear uno nuevo que importe la TBox).
  // Ahora mismo por cada fila de ASRS se genera un ejercicio/escenario independiente, con sus elementos dentro.    
  // Consejo: Sobre registrar la ejecución: https://developers.google.com/apps-script/guides/logging - Con Ctrl+Intro puedes ver lo que hay en el registro de ejecución. Entiendo que puede escribirse en el registro así: Logger.log(idEncabezados);  
  // Recorro toda la hoja de cálculo de entrada y voy contabilizando según lo que digan las traducciones
  

  var numIncidentes = matrizIncidentes.length;
  // Límites y contadores de las apariciones de cada palabra clave de Traducción    
  // - Los contenidos, tipos primitivos y propiedades, van en una hoja de cálculo de Traducciones. No usamos tipos definidos porque la gracia es que haya inferencia
  var numTraducciones = matrizTraducciones.length;
  var contadoresPalabrasClave = new Array(numTraducciones); 
  for (var numContador = 0; numContador < numTraducciones; numContador++) // No sé si es obligatorio, pero inicializo todos los contadores de palabras clave a 0  
    contadoresPalabrasClave[numContador] = 0; 

  /////////////////////////////////////////////////////////////////////////////////
  // Se monta el esquema de todos los individuos (escenarios)
  // Se podría crear también un TrainingSession que incluyera todos los escenarios, pero hemos preferido no hacerlo
  var infoIndividuos = plantillaEncabezado + plantillaEncabezadoIndividuos;
  
  // BUCLE PRINCIPAL (NIVEL 1): Recorremos indicente a incidente, informe a informe de la ASRS DB
  for (var numIncidente = 0; numIncidente < numIncidentes; numIncidente++) 
  {       
    var incidente = matrizIncidentes[numIncidente];
    
    // Este array de booleanos se utiliza por si hay algo que aparece en varias columnas, como Human Factors, para que sólo se contabilice una vez.
    var aparecidaPalabraClave = new Array(numTraducciones); 
    for (var numContador = 0; numContador < numTraducciones; numContador++) // No sé si es obligatorio, pero inicializo todos los booleanos de palabras clave a false
      aparecidaPalabraClave[numContador] = false; 
    
    /////////////////////////////////////////////////////////////////////////////////
    // Se monta el esquema de cada individuo (escenario)
    var infoIndividuo = plantillaIndividuo; 
    var atributos = ''; 
    var propiedades = ''; 
    var numSubindividuo = 0;
    var acn = 0; // Identificador ACN sin definir
    
    // BUCLE (NIVEL 2): Recorremos todas las columnas del incidente
    for (var numIdColumna = 0; numIdColumna < listaIdColumnas.length; numIdColumna++)
    {
      var idColumna = listaIdColumnas[numIdColumna];
      
      // BUCLE (NIVEL 3): Se recorren todas las filas de traducciones aplicables. ¡Lo de hacer este bucle dentro del bucle de las columnas es muy poco eficiente, la verdad!
      for (var numTraduccion = 0; numTraduccion < numTraducciones; numTraduccion++) 
      {         
        var palabraClave = matrizTraducciones[numTraduccion][0];
        var contenidoColumna = matrizTraducciones[numTraduccion][1];
        var idPropiedadoAtributo = matrizTraducciones[numTraduccion][2];
        var tipo = matrizTraducciones[numTraduccion][3];
  
        // Si se dan estas condiciones fundamentales... (debería poner un continue para los casos en que NO se cumplen)
        if (idColumna != undefined && idColumna.indexOf(contenidoColumna) > -1 && incidente[numIdColumna] != undefined) {
          // y si como palabra clave sale un asterisco (*) en la Traducción es que se trata de un atributo
          if (palabraClave == '*') {
            // Cojo la plantilla de atributo y cambio el atributo, poniendo además su tipo y buscando su valor para ponerlo también
            var atributo = plantillaAtributo;            
            atributo = atributo.replace(/\{\A\T\R\I\V\U\T\O\}/g, idPropiedadoAtributo);

            atributo = atributo.replace(/\{\T\I\P\O\_\A\T\R\I\V\U\T\O\}/g, tipo);
            atributo = atributo.replace(/\{\V\A\L\O\R\_\A\T\R\I\V\U\T\O\}/g, incidente[numIdColumna]);

            atributos+= atributo;

            // Si es el identificador ACN me lo guardo porque me sirve para nombrar al individuo
            if (contenidoColumna == 'ACN')
              acn = incidente[numIdColumna];

            // No estoy seguro de si en este caso hace falta llevar este conteo
            contadoresPalabrasClave[numTraduccion]++;
            aparecidaPalabraClave[numTraduccion] = true;    
          } else // En caso contrario, hay palabra clave.... y si además se dan estas otras de condiciones...

            if (!aparecidaPalabraClave[numTraduccion] && incidente[numIdColumna].indexOf(palabraClave) > -1) {                   
            
              // Cojo la plantilla de propiedad y cambio la propiedad, el ID y el NUM del padre... y luego el IDAUX y el NUMAUX del hijo
              var propiedad = plantillaPropiedad;
              propiedad = propiedad.replace(/\{\P\R\O\P\I\E\D\A\D\}/g, idPropiedadoAtributo);
                      
              /////////////////////////////////////////////////////////////////////////////////
              // También hay que crear los subindividuos (para cada propiedad). Los estamos agrupando como si fuesen parte del individuo raíz
              var infoSubindividuo = plantillaSubindividuo;
              var subTipos = plantillaTipo.replace(/\{\T\I\P\O\}/g, tipo); // Aquí sí podría haber varios tipos, ojo
              infoSubindividuo = infoSubindividuo.replace(/\{\T\I\P\O\S\}/g, subTipos);    
              // Los subindividuos, como siempre son elementos, se llamarán Element seguido del identificador ACN y un par de números (el de generación del escenario y el suyo propio) 
              // que van siempre incrementándose, empezando el primero por 00001 (formateado para que ocupe 5 dígitos) y el segundo por 01 (formateado para que ocupe 2 dígitos) 
              infoSubindividuo = infoSubindividuo.replace(/\{\I\D\}/g, 'Element' + acn);   

              var stringNum = '' + (numIncidente + 1); // Se convierte a string para poder añadirle ceros por delante
              while (stringNum.length < 5) 
                stringNum = '0' + stringNum;
              var stringSubnum = '' + (numSubindividuo + 1);  
              while (stringSubnum.length < 2) 
                stringSubnum = '0' + stringSubnum;
              infoSubindividuo = infoSubindividuo.replace(/\{\N\U\M\}/g, stringNum + "_" + stringSubnum);          
              
              infoIndividuos+= infoSubindividuo;
              numSubindividuo++;
              /////////////////////////////////////////////////////////////////////////////////
              
              // Otra vez nombramos al subindividuo
              propiedad = propiedad.replace(/\{\I\D\A\U\X}/g, 'Element' + acn);  
              propiedad = propiedad.replace(/\{\N\U\M\A\U\X}/g, stringNum + "_" + stringSubnum); 

              // En algunos casos (OWL Functional) hace falta nombrar al individuo padre cuando se está definiendo la propiedad
              propiedad = propiedad.replace(/\{\I\D\}/g, 'Scenario' + acn); 
              var stringNum = '' + (numIncidente + 1);  
              while (stringNum.length < 5) 
              stringNum = '0' + stringNum;
              propiedad = propiedad.replace(/\{\N\U\M\}/g, stringNum);

              propiedades+= propiedad;
              
              contadoresPalabrasClave[numTraduccion]++;
              aparecidaPalabraClave[numTraduccion] = true;    
            }
        }    

      } // BUCLE (NIVEL 3)     

    } // BUCLE (NIVEL 2)
    
    // Se rellena con las sustituciones debidas en el esquema del individuo 
    // - Sólo tiene un tipo (podrían ser varios tipos primitivos en un futuro... pero ya veremos)
    var tipos = plantillaTipo.replace(/\{\T\I\P\O\}/g, 'Scenario'); // En el individuo raíz sólo hay un tipo y es Scenario
    infoIndividuo = infoIndividuo.replace(/\{\T\I\P\O\S\}/g, tipos);  
    infoIndividuo = infoIndividuo.replace(/\{\A\T\R\I\V\U\T\O\S\}/g, atributos); // Habrá varios atributos      
    infoIndividuo = infoIndividuo.replace(/\{\P\R\O\P\I\E\D\A\D\E\S\}/g, propiedades);  // Aquí suele haber varias propiedades, ojo
    // Los individuos, como siempre son escenarios, se llamarán Scenario seguido del identificador ACN del informe, barra baja y finalmente un número 
    // (el de generación del escenario) que va siempre incrementándose, empezando el primero por 00001 (formateado para que ocupe 5 dígitos)  
    infoIndividuo = infoIndividuo.replace(/\{\I\D\}/g, 'Scenario' + acn);   
    var stringNum = '' + (numIncidente + 1);  
    while (stringNum.length < 5) 
      stringNum = '0' + stringNum;
    infoIndividuo = infoIndividuo.replace(/\{\N\U\M\}/g, stringNum);
    
    infoIndividuos+= infoIndividuo;    
  } // BUCLE PRINCIPAL (NIVEL 1)  

  infoIndividuos+= plantillaFinal;
  

  //////////////////////////////////////////////////////////////
  // Genero el nuevo informe en base a la plantilla y le pongo ID y fecha adecuados.
  
  // Google no permite escribir más de millón y pico de caracteres en un GoogleDocs, con lo que puede ser necesario generar VARIOS ficheros. Podemos ir partiendo cada medio millón, por ejemplo
  var idInformes;
  var maxCaracteres = 500000;
  var numFicheros = Math.ceil(infoIndividuos.length / maxCaracteres);
  
  // Cuando los generemos pondremos a todos los documentos generados la misma hora, basada en el uso horario de la sesión del que lo ejecuta, creo
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"); 
  var fechaCorta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd");
  var horaCorta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");  
  
  // Las estadísticas de incidentes tampoco tiene sentido hacerlas más de una vez, aunque por ahora sí las mostraremos en cada documento
  var infoIncidentes = '';

  // BUCLE POR TODAS LAS TRADUCCIONES
  for (var numTraduccion = 0; numTraduccion < numTraducciones; numTraduccion++) 
  { 
    infoIncidentes += 'Apariciones totales de ' + matrizTraducciones[numTraduccion][0] + 
      ' (' + matrizTraducciones[numTraduccion][1] + '): ' + 
        contadoresPalabrasClave[numTraduccion] + 
          ' (' + ((contadoresPalabrasClave[numTraduccion] / numIncidentes) * 100).toFixed(2) + '%)\n';
    
    // Mostrar también apariciones donde coincidan ciertas palabras clave (a la vez), e incluso con palabras clave de las de Human Factor
    // ...
  }
  
  // BUCLE POR TODOS LOS FICHEROS QUE VOY A GENERAR DE SALIDA
  for (var numFichero = 1; numFichero <= numFicheros; numFichero++)
  {     
    var idInforme = DriveApp.getFileById(idPlantillaInforme).makeCopy().getId(); 
 
    // Se podría incluir el nombre del fichero que se ha usado como entrada...
    DriveApp.getFileById(idInforme).setName(fechaCorta + ' ' + horaCorta + ' Salida autogenerada [' + numFichero + '/' + numFicheros + ']');
    
    // Abro el nuevo informe y sustituyo las plantillas entre llaves {} por la información adecuada
    var informe = DocumentApp.openById(idInforme);
    var cuerpoInforme = informe.getBody();  
    
    // Todo esto a lo mejor no debería repetirlo en todas las partes/ficheros que han salido
    cuerpoInforme.replaceText('{FECHA}', fecha); 
    cuerpoInforme.replaceText('{FILAS}', numIncidentes); 
    cuerpoInforme.replaceText('{INCIDENTES}', infoIncidentes);  
    
    //////////////////////////////////////////////////////////////
    // Completo la parte de los individuos del nuevo informe usando la información que he ido elaborando a partir de las filas de la hoja de cálculo de entrada
    // Ojo: Por ahora no estoy generando las relaciones en ambos sentidos, ni siendo muy exhaustivo en usar muchos tipos ni propiedades (no tengo ni los nombres bien puestos en la ontología...)
    // Ojo: Tampoco estoy añadiendo restricciones extra como que todos los individuos sean distintos entre sí, etc.
    
    // Sería más interesante cortar por líneas completas, y no en mitad de una línea, por caracteres, como hago ahora...
    cuerpoInforme.appendParagraph(infoIndividuos.substring((numFichero-1) * maxCaracteres, numFichero * maxCaracteres));  
  }
  
}