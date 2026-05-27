export function generateDiamondImgBase64(params: {
  threatLevel: string;
  threatName: string;
  personasAvg: number;
  personasLvl: string;
  recursosAvg: number;
  recursosLvl: string;
  sistemasAvg: number;
  sistemasLvl: string;
  redRombosCount: number;
}): string {
  // Crear un canvas en memoria
  const canvas = document.createElement('canvas');
  canvas.width = 400;
  canvas.height = 400;
  const ctx = canvas.getContext('2d');
  
  if (!ctx) return '';
  
  // Fondo
  ctx.fillStyle = '#1e293b';
  ctx.fillRect(0, 0, 400, 400);
  
  // Dibujar rombos (simplificado)
  ctx.fillStyle = params.threatLevel === 'INMINENTE' ? '#ef4444' : params.threatLevel === 'PROBABLE' ? '#f59e0b' : '#10b981';
  ctx.beginPath();
  ctx.moveTo(200, 40);
  ctx.lineTo(300, 120);
  ctx.lineTo(200, 200);
  ctx.lineTo(100, 120);
  ctx.fill();
  
  ctx.fillStyle = '#ffffff';
  ctx.font = 'bold 14px Arial';
  ctx.fillText('AMENAZA', 200, 110);
  
  // Retornar como data URL
  return canvas.toDataURL('image/png');
}