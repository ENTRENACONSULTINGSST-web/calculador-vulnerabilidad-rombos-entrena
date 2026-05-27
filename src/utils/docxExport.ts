import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';

export interface WordReportData {
  // Datos de la empresa (placeholders)
  empresaRazonSocial?: string;
  empresaNit?: string;
  empresaDireccion?: string;
  empresaTelefono?: string;
  empresaEmpleados?: number;
  empresaResponsableSST?: string;
  empresaRepresentanteLegal?: string;
  empresaCiudad?: string;
  
  // Datos de la amenaza
  threatName: string;
  threatDesc: string;
  threatLevel: string;
  
  // Resultados del diamante
  coreResults: any;
  transversalResults: any;
  
  // Respuestas del usuario
  answers: any;
  
  // Resultados de riesgo
  riskResult: any;
  vulLevel: string;
  redRombosCount: number;
  
  // Datos adicionales
  diamondDataURL?: string;
}

export async function exportToWord(data: WordReportData): Promise<void> {
  
  // ==================== FUNCIONES AUXILIARES ====================
  
  const getNivelTexto = (nivel: string): string => {
    switch(nivel) {
      case 'POSIBLE': return 'POSIBLE (Verde)';
      case 'PROBABLE': return 'PROBABLE (Amarillo)';
      case 'INMINENTE': return 'INMINENTE (Rojo)';
      default: return nivel;
    }
  };
  
  const getVulnerabilidadTexto = (nivel: string): string => {
    switch(nivel) {
      case 'BAJA': return 'BAJA (Verde) - Protegido';
      case 'MEDIA': return 'MEDIA (Amarillo) - Advertencia';
      case 'ALTA': return 'ALTA (Rojo) - Crítico';
      default: return nivel;
    }
  };
  
  // ==================== CONSTRUCCIÓN DEL DOCUMENTO ====================
  
  const children: any[] = [];
  
  // --- PORTADA / TÍTULO PRINCIPAL ---
  children.push(
    new Paragraph({
      text: "PLAN INSTITUCIONAL DE PREVENCIÓN Y EMERGENCIAS",
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "SG-SST - Sistema de Gestión de Seguridad y Salud en el Trabajo",
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );
  
  // --- 1. IDENTIFICACIÓN DE LA EMPRESA ---
  children.push(
    new Paragraph({
      text: "1. IDENTIFICACIÓN DE LA EMPRESA",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  const empresaData = [
    ["RAZÓN SOCIAL:", data.empresaRazonSocial || "[RAZÓN SOCIAL PENDIENTE]"],
    ["NIT:", data.empresaNit || "[NIT PENDIENTE]"],
    ["DIRECCIÓN:", data.empresaDireccion || "[DIRECCIÓN PENDIENTE]"],
    ["TELÉFONO:", data.empresaTelefono || "[TELÉFONO PENDIENTE]"],
    ["CIUDAD:", data.empresaCiudad || "[CIUDAD PENDIENTE]"],
    ["NÚMERO DE EMPLEADOS:", (data.empresaEmpleados || "[NÚMERO PENDIENTE]").toString()],
    ["RESPONSABLE SG-SST:", data.empresaResponsableSST || "[RESPONSABLE SG-SST PENDIENTE]"],
    ["REPRESENTANTE LEGAL:", data.empresaRepresentanteLegal || "[REPRESENTANTE LEGAL PENDIENTE]"],
  ];
  
  const empresaTableRows = empresaData.map(row => 
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(row[0])], width: { size: 30, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph(row[1])], width: { size: 70, type: WidthType.PERCENTAGE } }),
      ],
    })
  );
  
  children.push(
    new Table({
      rows: empresaTableRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    })
  );
  
  // --- 2. MARCO NORMATIVO (texto fijo) ---
  children.push(
    new Paragraph({
      text: "2. MARCO NORMATIVO",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph("La legislación colombiana en materia de salud ocupacional establece en varias normas la obligatoriedad que tienen las entidades públicas y privadas para implementar el Programa Integral para la Prevención y el Control de Emergencias, todas fundamentadas en la obligación de todos los empleadores de '...garantizar la salud de los trabajadores...' (Numeral 348 del Código Sustantivo del Trabajo)."),
    new Paragraph({ text: "", spacing: { after: 200 } }),
    new Paragraph("Decreto 1072 de 2015 - Artículo 2.2.4.6.25: establece la obligación de implementar un Plan de Prevención, Preparación y Respuesta ante Emergencias."),
    new Paragraph("Resolución 0312 de 2019 - Define los Estándares Mínimos del SG-SST, incluyendo plan de emergencias, brigada de emergencias, capacitación, simulacros e inspección de equipos."),
    new Paragraph("Ley 1523 de 2012 - Adopta la Política Nacional de Gestión del Riesgo de Desastres."),
  );
  
  // --- 3. AMENAZAS IDENTIFICADAS ---
  children.push(
    new Paragraph({
      text: "3. AMENAZAS IDENTIFICADAS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  const amenazasList = [
    { nombre: "Incendio Estructural / Explosión", nivel: "POSIBLE", activa: true },
    { nombre: "Sismo / Terremoto", nivel: "PROBABLE", activa: true },
    { nombre: "Inundación por Lluvias / Escorrentía", nivel: "POSIBLE", activa: true },
    { nombre: "Derrame de Materiales Peligrosos", nivel: "POSIBLE", activa: true },
    { nombre: "Hurto / Asonada / Terrorismo (Riesgo Público)", nivel: "POSIBLE", activa: true },
    { nombre: "Embalamiento Térmico de Baterías de Litio", nivel: "POSIBLE", activa: true },
    { nombre: "Trabajo en Alturas", nivel: "POSIBLE", activa: true },
    { nombre: "Riesgo Psicosocial", nivel: "POSIBLE", activa: true },
  ];
  
  const amenazasRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "AMENAZA", bold: true })], width: { size: 50, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "NIVEL DE PROBABILIDAD", bold: true })], width: { size: 30, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "ESTADO", bold: true })], width: { size: 20, type: WidthType.PERCENTAGE } }),
      ],
    }),
  ];
  
  amenazasList.forEach((amenaza) => {
    amenazasRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(amenaza.nombre)] }),
          new TableCell({ children: [new Paragraph(getNivelTexto(amenaza.nivel))] }),
          new TableCell({ children: [new Paragraph(amenaza.activa ? "ACTIVA" : "INACTIVA")] }),
        ],
      })
    );
  });
  
  children.push(
    new Table({
      rows: amenazasRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    })
  );
  
  // --- 4. ANÁLISIS DE VULNERABILIDAD (con ítems detallados) ---
  children.push(
    new Paragraph({
      text: "4. ANÁLISIS DE VULNERABILIDAD",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  // Definir los ítems para cada componente
  const itemsPorComponente: { [key: string]: { id: string; name: string; questions: string[] }[] } = {
    Personas: [
      { id: "P_Organizacion", name: "Organización", questions: ["p_org_q1", "p_org_q2", "p_org_q3"] },
      { id: "P_Capacitacion", name: "Capacitación", questions: ["p_cap_q1", "p_cap_q2", "p_cap_q3"] },
      { id: "P_Dotacion", name: "Dotación", questions: ["p_dot_q1", "p_dot_q2", "p_dot_q3"] },
    ],
    Recursos: [
      { id: "R_Materiales", name: "Materiales", questions: ["r_mat_q1", "r_mat_q2", "r_mat_q3"] },
      { id: "R_Infraestructura", name: "Infraestructura", questions: ["r_inf_q1", "r_inf_q2", "r_inf_q3"] },
      { id: "R_Equipos", name: "Equipos", questions: ["r_equ_q1", "r_equ_q2", "r_equ_q3"] },
    ],
    Sistemas: [
      { id: "S_Servicios", name: "Servicios Públicos", questions: ["s_ser_q1", "s_ser_q2", "s_ser_q3"] },
      { id: "S_Alternos", name: "Sistemas Alternos", questions: ["s_alt_q1", "s_alt_q2", "s_alt_q3"] },
      { id: "S_Recuperacion", name: "Recuperación", questions: ["s_rec_q1", "s_rec_q2", "s_rec_q3"] },
    ],
  };
  
  // Calcular promedios por ítem
  const calcularPromedioItem = (questionIds: string[]): number => {
    let sum = 0;
    let count = 0;
    questionIds.forEach(id => {
      if (data.answers[id] !== undefined) {
        sum += data.answers[id];
        count++;
      }
    });
    return count > 0 ? parseFloat((sum / count).toFixed(2)) : 0;
  };
  
  // Generar tabla para cada componente
  const componentes = ["Personas", "Recursos", "Sistemas"];
  
  componentes.forEach(componente => {
    children.push(
      new Paragraph({
        text: `4.${componentes.indexOf(componente) + 1} ${componente}`,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 100 },
      })
    );
    
    const itemsRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: "ÍTEM", bold: true })], width: { size: 40, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [new Paragraph({ text: "PROMEDIO", bold: true })], width: { size: 30, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [new Paragraph({ text: "CALIFICACIÓN", bold: true })], width: { size: 30, type: WidthType.PERCENTAGE } }),
        ],
      }),
    ];
    
    const items = itemsPorComponente[componente];
    items.forEach(item => {
      const avg = calcularPromedioItem(item.questions);
      let nivel = "";
      if (avg <= 0.40) nivel = "ALTA (Crítico)";
      else if (avg <= 0.60) nivel = "MEDIA (Advertencia)";
      else nivel = "BAJA (Protegido)";
      
      itemsRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(item.name)] }),
            new TableCell({ children: [new Paragraph(avg.toFixed(2))] }),
            new TableCell({ children: [new Paragraph(nivel)] }),
          ],
        })
      );
    });
    
    children.push(
      new Table({
        rows: itemsRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1 },
          bottom: { style: BorderStyle.SINGLE, size: 1 },
          left: { style: BorderStyle.SINGLE, size: 1 },
          right: { style: BorderStyle.SINGLE, size: 1 },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
          insideVertical: { style: BorderStyle.SINGLE, size: 1 },
        },
      })
    );
  });
  
  // --- 5. NIVEL DE RIESGO CONSOLIDADO ---
  children.push(
    new Paragraph({
      text: "5. NIVEL DE RIESGO CONSOLIDADO",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph(`Vulnerabilidad Consolidada: ${getVulnerabilidadTexto(data.vulLevel)}`),
    new Paragraph(`Rombos en Rojo: ${data.redRombosCount} de 3 componentes`),
    new Paragraph(`Riesgo Global: ${data.riskResult?.level || "MEDIO"}`),
    new Paragraph(data.riskResult?.description || ""),
    new Paragraph({ text: "", spacing: { after: 200 } })
  );
  
  // --- 6. RECURSOS PARA EMERGENCIAS ---
  children.push(
    new Paragraph({
      text: "6. RECURSOS PARA EMERGENCIAS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph({ text: "6.1 Recursos de infraestructura y equipos", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } })
  );
  
  const equiposRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "DESCRIPCIÓN", bold: true })], width: { size: 50, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "CANTIDAD", bold: true })], width: { size: 20, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "UBICACIÓN", bold: true })], width: { size: 30, type: WidthType.PERCENTAGE } }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph("[DESCRIPCIÓN DEL EQUIPO]")] }),
        new TableCell({ children: [new Paragraph("[CANTIDAD]")] }),
        new TableCell({ children: [new Paragraph("[UBICACIÓN]")] }),
      ],
    }),
  ];
  
  children.push(
    new Table({
      rows: equiposRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    }),
    new Paragraph({ text: "6.2 Recursos externos", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } })
  );
  
  const externosRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "ENTIDAD", bold: true })], width: { size: 40, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "CLASE DE AYUDA", bold: true })], width: { size: 35, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "TELÉFONO", bold: true })], width: { size: 25, type: WidthType.PERCENTAGE } }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph("[ENTIDAD]")] }),
        new TableCell({ children: [new Paragraph("[CLASE DE AYUDA]")] }),
        new TableCell({ children: [new Paragraph("[TELÉFONO]")] }),
      ],
    }),
  ];
  
  children.push(
    new Table({
      rows: externosRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    })
  );
  
  // --- 7. PROCEDIMIENTOS OPERATIVOS NORMALIZADOS (PON) ---
  children.push(
    new Paragraph({
      text: "7. PROCEDIMIENTOS OPERATIVOS NORMALIZADOS (PON)",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph({ text: "7.1 PON PARA ACTUAR EN CASO DE INCENDIO", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
    new Paragraph("1. Detección del fuego / Activación de alarma manual"),
    new Paragraph("2. Evaluar si el fuego es un conato (controlable con extintor)"),
    new Paragraph("3. Si es conato: desplegar extintor portátil adecuado (PQS/CO2) a la base del fuego"),
    new Paragraph("4. Si el fuego está fuera de control: activar alarma general e iniciar evacuación"),
    new Paragraph("5. Corte inmediato de acometida general de energía y gas"),
    new Paragraph("6. Llamar a Bomberos Oficiales (123) y dar ubicación exacta"),
    new Paragraph("7. Guiar al personal al punto de encuentro y realizar conteo"),
    new Paragraph("8. Entregar el mando a Bomberos y evaluar el retorno seguro"),
    new Paragraph({ text: "", spacing: { after: 200 } }),
    new Paragraph({ text: "7.2 PON PARA EVACUACIÓN", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
    new Paragraph("1. Al escuchar la alarma de evacuación, conserve la calma"),
    new Paragraph("2. Suspenda inmediatamente su actividad laboral"),
    new Paragraph("3. Siga las instrucciones del coordinador de evacuación y los brigadistas"),
    new Paragraph("4. Diríjase hacia la ruta de evacuación señalizada más cercana"),
    new Paragraph("5. No corra, no grite, no empuje a otros"),
    new Paragraph("6. Si hay humo, desplácese agachado y cubriendo nariz y boca"),
    new Paragraph("7. Diríjase al punto de encuentro establecido"),
    new Paragraph("8. Espere instrucciones del coordinador de evacuación"),
    new Paragraph({ text: "", spacing: { after: 200 } }),
    new Paragraph({ text: "7.3 PON PARA SISMO", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
    new Paragraph("1. Durante el sismo: NO evacuar. Adopte posición de autoprotección"),
    new Paragraph("2. Ubíquese bajo vigas robustas, escritorios o alejado de ventanas"),
    new Paragraph("3. Al cesar el movimiento: evalúe salida y daños parciales"),
    new Paragraph("4. Ordene evacuación inmediata por rutas preestablecidas"),
    new Paragraph("5. Revise redes de gas y electricidad en el exterior"),
    new Paragraph("6. Concéntrese en el punto de encuentro y atienda heridos"),
  );
  
  // --- 8. FIRMAS Y APROBACIONES ---
  children.push(
    new Paragraph({
      text: "8. FIRMAS Y APROBACIONES",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph({ text: "", spacing: { after: 400 } }),
    new Paragraph({ text: "_________________________________________", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: data.empresaResponsableSST || "[RESPONSABLE SG-SST PENDIENTE]", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "Responsable del SG-SST", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "", spacing: { after: 400 } }),
    new Paragraph({ text: "_________________________________________", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: data.empresaRepresentanteLegal || "[REPRESENTANTE LEGAL PENDIENTE]", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "Representante Legal", alignment: AlignmentType.CENTER }),
  );
  
  // --- 9. CONTROL DE CAMBIOS ---
  children.push(
    new Paragraph({
      text: "9. CONTROL DE CAMBIOS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  const cambiosRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "FECHA", bold: true })], width: { size: 20, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "VERSIÓN", bold: true })], width: { size: 15, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "MOTIVO DE LA MODIFICACIÓN", bold: true })], width: { size: 45, type: WidthType.PERCENTAGE } }),
        new TableCell({ children: [new Paragraph({ text: "RESPONSABLE", bold: true })], width: { size: 20, type: WidthType.PERCENTAGE } }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(new Date().toLocaleDateString('es-CO'))] }),
        new TableCell({ children: [new Paragraph("1.0")] }),
        new TableCell({ children: [new Paragraph("Creación del plan de emergencias según metodología Diamante de Riesgo")] }),
        new TableCell({ children: [new Paragraph(data.empresaResponsableSST || "[RESPONSABLE SG-SST PENDIENTE]")] }),
      ],
    }),
  ];
  
  children.push(
    new Table({
      rows: cambiosRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    })
  );
  
  // ==================== GENERAR Y DESCARGAR ====================
  
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: children,
    }],
  });
  
  const blob = await Packer.toBlob(doc);
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Plan_Emergencias_${data.threatName.replace(/\s/g, '_')}_${new Date().toLocaleDateString('es-CO').replace(/\//g, '-')}.docx`;
  link.click();
  URL.revokeObjectURL(link.href);
}