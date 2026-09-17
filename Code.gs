// Escrito de Teoría de conjuntos — Code.gs
// Pegar en Apps Script e Implementar → Web App

var SHEET_NAME = "Respuestas";

var HEADERS = [
  "Timestamp",
  "Nombre",
  "P1 – Elija la fórmula que corresponda con la parte blanca de la siguiente imagen: ",
  "P2 – Arrastre las expresiones que correspondan para que cada conjunto quede expresado por comprensión",
  "P3 – Elija la fórmula que corresponda con la parte gris clara de la siguiente imagen:",
  "P4 – Existen algunos conjuntos particulares, por ejemplo los conjuntos [...] son aquellos que tienen un solo elemento, en cam",
  "P5 – Sea U={1,2,3,4,5,6,7,8,9,10} el conjunto universal, A={1,4,7,10}, B={1,2,3,4,5} y C={2,4,6,8} conjuntos en U. Definir po",
  "P6 – De acuerdo al siguiente diagrama, marque las opciones correctas:",
  "P7 – Dado los siguientes conjuntos A={a,b,c,d}, B={c,d,e,f}, C={f,g,h} y U={a,b,c,d,e,f,g,h,i}. Arrastre la opción correcta a",
  "Resultado (pts)"
];

var QUESTION_IDS = ["q1","q2","q3","q4","q5","q6","q7"];

var CORRECT_ANSWERS = {
  "q1": "",
  "q3": ""
};

var CORRECT_MULTI = {
  "q6": [
    "c",
    "d"
  ]
};

var CORRECT_PAIRS = {
  "q2": [
    [
      "A = {4, 6, 8, 10}",
      "{x / x ∈ ℕ, 3 ≤ x ≤ 10, x = mult(2)}"
    ],
    [
      "B = {3, 5, 7, 9}",
      "{x / x ∈ ℕ, 3 ≤ x ≤ 10, x ≠ mult(2)}"
    ],
    [
      "C = {3, 4, ..., 9, 10}",
      "{x / x ∈ ℕ, 3 ≤ x ≤ 10}"
    ]
  ],
  "q5": [
    [
      "B̄ ∩ (C − A)",
      "{6, 8}"
    ],
    [
      "(A ∩ B)̄ ∪ C",
      "{2, 3, 4, 5, 6, 7, 8, 9, 10}"
    ],
    [
      "A ⊕ C",
      "{3, 4, 5, 9}"
    ]
  ],
  "q7": [
    [
      "B − A",
      "{e, f}"
    ],
    [
      "Cᶜ",
      "{a, b, c, d, e, i}"
    ],
    [
      "(A ∩ C) ∪ B",
      "{c, d, e, f}"
    ]
  ]
};

var CORRECT_FILLS = {
  "q4": [
    "unitarios",
    "vacíos",
    "finitos",
    "infinitos"
  ]
};

var QUESTION_POINTS = {
  "q1": 1,
  "q2": 2,
  "q3": 1,
  "q4": 2,
  "q5": 2,
  "q6": 2,
  "q7": 2
};

function doGet(e){
  var cb=e.parameter.callback||"callback";
  try{
  var data={};
  try{data=JSON.parse(e.parameter.data||"{}");}catch(err){}

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(SHEET_NAME);
  if(!sheet){sheet=ss.insertSheet(SHEET_NAME);sheet.appendRow(HEADERS);}

  var earned=0,totalPts=0,correct=0,results={};
  QUESTION_IDS.forEach(function(id){
    var p=QUESTION_POINTS[id]||1;
    if(CORRECT_ANSWERS[id]!==undefined){
      totalPts+=p;
      var ok=String(data[id]||"").trim().toLowerCase()===String(CORRECT_ANSWERS[id]).trim().toLowerCase();
      if(ok){earned+=p;correct++;}
      results[id]=ok?"correcta":"incorrecta";
    } else if(CORRECT_MULTI[id]){
      totalPts+=p;
      var subM=(data[id]||"").split(",").map(function(s){return s.trim();}).filter(Boolean).sort().join(",");
      var expM=(CORRECT_MULTI[id]||[]).slice().sort().join(",");
      var okM=subM===expM;
      if(okM){earned+=p;correct++;}
      results[id]=okM?"correcta":"incorrecta";
    } else if(CORRECT_PAIRS[id]){
      totalPts+=p;
      var submitted={};
      try{submitted=JSON.parse(data[id]||"{}");}catch(e2){}
      var pairList=CORRECT_PAIRS[id], got=0;
      pairList.forEach(function(pair){ if(submitted[pair[0]]===pair[1]) got++; });
      var ratio=pairList.length>0?got/pairList.length:0;
      earned+=p*ratio;
      if(ratio===1)correct++;
      results[id]=ratio===1?"correcta":"parcial:"+got+"/"+pairList.length;
    } else if(CORRECT_FILLS[id]){
      totalPts+=p;
      var submittedFills=(data[id]||"").split("|||");
      var fillList=CORRECT_FILLS[id], fillGot=0;
      fillList.forEach(function(ans,fi){ if((submittedFills[fi]||"").trim().toLowerCase()===ans.trim().toLowerCase()) fillGot++; });
      var fillRatio=fillList.length>0?fillGot/fillList.length:0;
      earned+=p*fillRatio;
      if(fillRatio===1)correct++;
      results[id]=fillRatio===1?"correcta":"parcial:"+fillGot+"/"+fillList.length;
    } else { results[id]="abierta"; }
  });

  var row=[new Date(),data.nombre||""];
  QUESTION_IDS.forEach(function(id){row.push(data[id]||"");});
  row.push(earned.toFixed(2)+"/"+totalPts);
  sheet.appendRow(row);

  var payload={ok:true,correct:correct,total:QUESTION_IDS.length,earned:earned,totalPts:totalPts,results:results};
  return ContentService
    .createTextOutput(cb+"("+JSON.stringify(payload)+")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch(ex){
    return ContentService
      .createTextOutput(cb+'({"ok":false,"error":'+JSON.stringify(String(ex))+'})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}
