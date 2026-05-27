import * as XLSX from 'xlsx';

export interface ExcelReportData {
  threatName: string;
  threatDesc: string;
  threatLevel: string;
  coreResults: any;
  transversalResults: any;
  answers: any;
  riskResult: any;
  vulLevel: string;
  redRombosCount: number;
}

export function exportToExcel(data: ExcelReportData): void {
  // Crear libro de trabajo
  const workbook = XLSX.utils.book_new();
  
  // Hoja 1: Resumen General
  const summaryData = [
    ['INFORME DE VULNERABILIDAD POR ROMBOS'],
    ['ENTRENA CONSULTING SAS'],
    [''],
    ['Amenaza Evaluada:', data.threatName],
    ['Nivel de Probabilidad:', data.threatLevel],
    ['Descripción:', data.threatDesc],
    [''],
    ['Vulnerabilidad Consolidada:', data.vulLevel],
    ['Rombos en Rojo:', data.redRombosCount.toString()],
    ['Riesgo Global:', data.riskResult.level],
    ['Descripción del Riesgo:', data.riskResult.description],
    [''],
    ['COMPONENTES CORE'],
    ['Componente', 'Promedio', 'Nivel'],
    ['Personas', data.coreResults?.Personas?.average?.toString() || '0', data.coreResults?.Personas?.level || 'No definido'],
    ['Recursos', data.coreResults?.Recursos?.average?.toString() || '0', data.coreResults?.Recursos?.level || 'No definido'],
    ['Sistemas', data.coreResults?.Sistemas?.average?.toString() || '0', data.coreResults?.Sistemas?.level || 'No definido'],
  ];
  
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumen General');
  
  // Hoja 2: Módulos Transversales
  const transversalData = [
    ['MÓDULOS TRANSVERSALES'],
    ['Bloque', 'Promedio', 'Nivel', 'Referencia Legal'],
  ];
  
  if (data.transversalResults) {
    Object.values(data.transversalResults).forEach((t: any) => {
      transversalData.push([t.name, t.average?.toString() || '0', t.level || 'No definido', t.lawReference || '']);
    });
  }
  
  const transversalSheet = XLSX.utils.aoa_to_sheet(transversalData);
  XLSX.utils.book_append_sheet(workbook, transversalSheet, 'Módulos Transversales');
  
  // Generar archivo
  XLSX.writeFile(workbook, `Memoria_Tecnica_${data.threatName.replace(/\s/g, '_')}.xlsx`);
}