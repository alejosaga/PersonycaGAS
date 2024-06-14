function cotizar() {
  
  addRowNumber(SSmaestroCot,"Datos",lastColumnDat-3);
    
  switch(servicio) {
    case "Consultoria SG-Seguridad y salud en el trabajo":
      if(claseRiesgo<4 && numEmp == "Menos de 11"){
      sieteEstandares() 
      }
      else{
      sgSst();
      }
      break;
    case "Compliance":
      //function2();
      break;
    case "Aplicacion Bateria riesgo psicosocial":
      batPsiRisk();
      break;
    case "Programa Etica empresarial":
      //function3();
      break;
    case "Sistemas de Gestion de Calidad -ISO":
      generarCotizacionISO();
      break;
    case "Fortalecimiento Talento Humano":
      //function3();
      break;
    case "Sitemas integrados de gestion":
      //function3();
      break;
    case "Gestion de Calidad":
      //function3();
      break;
    default:
      // Default actions
      break;
  }
  
}
