import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx';

export interface WordReportData {
  // Datos de la empresa (desde la interfaz)
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
  
  // Amenazas personalizadas (desde el gestor)
  amenazasPersonalizadas?: any[];
  
  // Recursos (equipos y externos)
  recursosEquipos?: any[];
  recursosExternos?: any[];
  
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
  
  // --- TÍTULO PRINCIPAL ---
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
  
  // --- TABLA DE DATOS DE LA EMPRESA ---
  children.push(
    new Paragraph({
      text: "1. IDENTIFICACIÓN DE LA EMPRESA",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 200, after: 200 },
    })
  );
  
  const empresaData = [
    ["RAZÓN SOCIAL:", data.empresaRazonSocial || "[RAZÓN SOCIAL]"],
    ["NIT:", data.empresaNit || "[NIT]"],
    ["DIRECCIÓN:", data.empresaDireccion || "[DIRECCIÓN]"],
    ["TELÉFONO:", data.empresaTelefono || "[TELÉFONO]"],
    ["CIUDAD:", data.empresaCiudad || "[CIUDAD]"],
    ["NÚMERO DE EMPLEADOS:", (data.empresaEmpleados || "[NÚMERO]").toString()],
    ["RESPONSABLE SG-SST:", data.empresaResponsableSST || "[RESPONSABLE SG-SST]"],
    ["REPRESENTANTE LEGAL:", data.empresaRepresentanteLegal || "[REPRESENTANTE LEGAL]"],
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
  
  // --- AMENAZAS IDENTIFICADAS ---
  children.push(
    new Paragraph({
      text: "2. AMENAZAS IDENTIFICADAS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  // Usar amenazas personalizadas o las estándar
  const amenazasList = data.amenazasPersonalizadas && data.amenazasPersonalizadas.length > 0 
    ? data.amenazasPersonalizadas 
    : [
        { nombre: "Incendio Estructural / Explosión", nivel: "POSIBLE", activa: true },
        { nombre: "Sismo / Terremoto", nivel: "PROBABLE", activa: true },
        { nombre: "Inundación por Lluvias", nivel: "POSIBLE", activa: true },
        { nombre: "Derrame de Materiales Peligrosos", nivel: "POSIBLE", activa: true },
        { nombre: "Hurto / Asonada / Terrorismo", nivel: "POSIBLE", activa: true },
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
  
  amenazasList.forEach((amenaza: any) => {
    amenazasRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(amenaza.nombre || amenaza.name || amenaza.threatName || "")] }),
          new TableCell({ children: [new Paragraph(getNivelTexto(amenaza.nivel || amenaza.level || amenaza.threatLevel || "POSIBLE"))] }),
          new TableCell({ children: [new Paragraph(amenaza.activa !== false ? "ACTIVA" : "INACTIVA")] }),
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
  
  // --- ANÁLISIS DE VULNERABILIDAD CORE ---
  children.push(
    new Paragraph({
      text: "3. ANÁLISIS DE VULNERABILIDAD CORE",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  // Tabla de Personas
  if (data.coreResults?.Personas) {
    const p = data.coreResults.Personas;
    children.push(
      new Paragraph({ text: "3.1 Personas", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
      new Paragraph(`Promedio: ${p.average.toFixed(2)} - ${getVulnerabilidadTexto(p.level)}`),
      new Paragraph({ text: "", spacing: { after: 200 } })
    );
  }
  
  // Tabla de Recursos
  if (data.coreResults?.Recursos) {
    const r = data.coreResults.Recursos;
    children.push(
      new Paragraph({ text: "3.2 Recursos", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
      new Paragraph(`Promedio: ${r.average.toFixed(2)} - ${getVulnerabilidadTexto(r.level)}`),
      new Paragraph({ text: "", spacing: { after: 200 } })
    );
  }
  
  // Tabla de Sistemas
  if (data.coreResults?.Sistemas) {
    const s = data.coreResults.Sistemas;
    children.push(
      new Paragraph({ text: "3.3 Sistemas y Procesos", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }),
      new Paragraph(`Promedio: ${s.average.toFixed(2)} - ${getVulnerabilidadTexto(s.level)}`),
      new Paragraph({ text: "", spacing: { after: 200 } })
    );
  }
  
  // --- NIVEL DE RIESGO CONSOLIDADO ---
  children.push(
    new Paragraph({
      text: "4. NIVEL DE RIESGO CONSOLIDADO",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph(`Vulnerabilidad Consolidada: ${getVulnerabilidadTexto(data.vulLevel)}`),
    new Paragraph(`Rombos en Rojo: ${data.redRombosCount} de 3 componentes`),
    new Paragraph(`Riesgo Global: ${data.riskResult?.level || "MEDIO"}`),
    new Paragraph(data.riskResult?.description || ""),
    new Paragraph({ text: "", spacing: { after: 200 } })
  );
  
  // --- MÓDULOS TRANSVERSALES (opcional) ---
  if (data.transversalResults && Object.keys(data.transversalResults).length > 0) {
    children.push(
      new Paragraph({
        text: "5. MÓDULOS TRANSVERSALES ESPECIALES",
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 200 },
      })
    );
    
    const transRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: "MÓDULO", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "PROMEDIO", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "CALIFICACIÓN", bold: true })] }),
        ],
      }),
    ];
    
    Object.values(data.transversalResults).forEach((t: any) => {
      transRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(t.name)] }),
            new TableCell({ children: [new Paragraph(t.average?.toFixed(2) || "0.00")] }),
            new TableCell({ children: [new Paragraph(getVulnerabilidadTexto(t.level))] }),
          ],
        })
      );
    });
    
    children.push(
      new Table({
        rows: transRows,
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
  }
  
  // --- RECURSOS PARA EMERGENCIAS ---
  children.push(
    new Paragraph({
      text: "6. RECURSOS PARA EMERGENCIAS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    })
  );
  
  // Equipos
  if (data.recursosEquipos && data.recursosEquipos.length > 0) {
    children.push(new Paragraph({ text: "6.1 Equipos e Infraestructura", heading: HeadingLevel.HEADING_2, spacing: { after: 100 } }));
    
    const equiposRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: "DESCRIPCIÓN", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "CANTIDAD", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "UBICACIÓN", bold: true })] }),
        ],
      }),
    ];
    
    data.recursosEquipos.forEach((e: any) => {
      equiposRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(e.descripcion || "")] }),
            new TableCell({ children: [new Paragraph((e.cantidad || "0").toString())] }),
            new TableCell({ children: [new Paragraph(e.ubicacion || "")] }),
          ],
        })
      );
    });
    
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
      })
    );
  }
  
  // Contactos externos
  if (data.recursosExternos && data.recursosExternos.length > 0) {
    children.push(new Paragraph({ text: "6.2 Contactos Externos de Emergencia", heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
    
    const externosRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: "ENTIDAD", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "CLASE DE AYUDA", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "TELÉFONO", bold: true })] }),
        ],
      }),
    ];
    
    data.recursosExternos.forEach((c: any) => {
      externosRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(c.entidad || "")] }),
            new TableCell({ children: [new Paragraph(c.ayuda || c.claseAyuda || "")] }),
            new TableCell({ children: [new Paragraph(c.telefono || "")] }),
          ],
        })
      );
    });
    
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
  }
  
  // --- HALLAZGOS Y RECOMENDACIONES ---
  children.push(
    new Paragraph({
      text: "7. HALLAZGOS Y RECOMENDACIONES",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph("A continuación, se listan las preguntas calificadas como NO (0.0) o PARCIAL (0.5) durante la evaluación, junto con las recomendaciones técnicas para su corrección:")
  );
  
  if (data.answers) {
    const hallazgosRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: "PREGUNTA", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "CALIFICACIÓN", bold: true })] }),
          new TableCell({ children: [new Paragraph({ text: "RECOMENDACIÓN", bold: true })] }),
        ],
      }),
    ];
    
    Object.entries(data.answers).forEach(([id, value]: [string, any]) => {
      if (value === 0.0 || value === 0.5) {
        hallazgosRows.push(
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph(id)] }),
              new TableCell({ children: [new Paragraph(value === 0.0 ? "NO (0.0)" : "PARCIAL (0.5)")] }),
              new TableCell({ children: [new Paragraph("Revisar y corregir según normativa aplicable")] }),
            ],
          })
        );
      }
    });
    
    children.push(
      new Table({
        rows: hallazgosRows,
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
  }
  
  // --- FIRMAS Y CONTROL DE CAMBIOS ---
  children.push(
    new Paragraph({
      text: "8. FIRMAS Y APROBACIONES",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }),
    new Paragraph({ text: "", spacing: { after: 400 } }),
    new Paragraph({ text: "_________________________________________", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: data.empresaResponsableSST || "[RESPONSABLE SG-SST]", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "Responsable del SG-SST", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "", spacing: { after: 400 } }),
    new Paragraph({ text: "_________________________________________", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: data.empresaRepresentanteLegal || "[REPRESENTANTE LEGAL]", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "Representante Legal", alignment: AlignmentType.CENTER }),
    new Paragraph({ text: "", spacing: { after: 400 } }),
    new Paragraph({
      text: "9. CONTROL DE CAMBIOS",
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 200, after: 200 },
    })
  );
  
  const cambiosRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "FECHA", bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: "VERSIÓN", bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: "MOTIVO DE LA MODIFICACIÓN", bold: true })] }),
        new TableCell({ children: [new Paragraph({ text: "RESPONSABLE", bold: true })] }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(new Date().toLocaleDateString('es-CO'))] }),
        new TableCell({ children: [new Paragraph("1.0")] }),
        new TableCell({ children: [new Paragraph("Creación del plan de emergencias según metodología Diamante de Riesgo")] }),
        new TableCell({ children: [new Paragraph(data.empresaResponsableSST || "[RESPONSABLE SG-SST]")] }),
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
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1 pulgada en twips
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