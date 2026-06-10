// ============================================================
// NEUROFEEDBACK ILF - Google Apps Script v7
// Correspondances exactes d'après le tableau original
// 1=certain(X), 2=incertain(?)
// ============================================================

function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("processNewEmails").timeBased().everyMinutes(5).create();
  Logger.log("Declencheur cree !");
}

function processNewEmails() {
  var label = "neurofeedback-traite";
  var processedLabel = GmailApp.getUserLabelByName(label);
  if (!processedLabel) processedLabel = GmailApp.createLabel(label);
  var threads = GmailApp.search("subject:\"Questionnaire Neurofeedback\" -label:" + label);
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      var body = message.getPlainBody();
      if (body && (body.indexOf("Nom :") !== -1 || body.indexOf("Patient :") !== -1)) {
        parseAndSave(body, message.getDate());
      }
    });
    thread.addLabel(processedLabel);
  });
}

function extractField(body, key) {
  var idx = body.indexOf(key);
  if (idx === -1) return "";
  var start = idx + key.length;
  while (start < body.length && body[start] === " ") start++;
  var nextLine = body.indexOf("\n", start);
  var val = nextLine !== -1 ? body.substring(start, nextLine) : body.substring(start);
  return val.trim();
}

function extractPatientName(body) {
  var idx = body.indexOf("Patient :");
  if (idx !== -1) {
    var start = idx + 9;
    var nextLine = body.indexOf("\n", start);
    return (nextLine !== -1 ? body.substring(start, nextLine) : body.substring(start)).trim();
  }
  var nom = extractField(body, "Nom :");
  var prenom = extractField(body, "Prénom :");
  return (prenom + " " + nom).trim();
}

function parseAndSave(body, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullName = extractPatientName(body);
  var nameParts = fullName.split(" ");
  var data = {
    nom: nameParts.slice(1).join(" ") || fullName,
    prenom: nameParts[0] || "",
    age: extractField(body, "Age :"),
    ddn: extractField(body, "Date naissance :"),
    email: extractField(body, "Email :"),
    tel: extractField(body, "Téléphone :"),
    ville: extractField(body, "Ville :"),
    main: extractField(body, "Main :"),
    raison: extractField(body, "Raison :")
  };
  var answers = [];
  var traumaSet = {}; // qnums where "Si trauma" is checked by practitioner
  var lines = body.split("\n");
  lines.forEach(function(line) {
    line = line.trim();
    // Detect practitioner flags: "Q3 [praticien] : Si trauma" ou "Q20 [praticien] : Attention aux détails"
    if (line.indexOf("[praticien]") !== -1) {
      var colonIdx2 = line.indexOf(" : ");
      if (colonIdx2 !== -1) {
        var qnum2 = parseInt(line.substring(1));
        var condLabel = line.substring(colonIdx2 + 3).trim().toLowerCase();
        if (!isNaN(qnum2)) {
          if (condLabel === "si trauma") {
            traumaSet[qnum2] = true; // backward compat
          }
          traumaSet[qnum2 + "_" + condLabel] = true;
        }
      }
    }
    if (line.indexOf("Q") === 0) {
      var spaceIdx = line.indexOf(" ");
      if (spaceIdx === -1) return;
      var qnum = parseInt(line.substring(1, spaceIdx));
      if (isNaN(qnum) || qnum < 1 || qnum > 300) return;
      var rest = line.substring(spaceIdx + 1).replace(/^[–—] /, "");
      var colonIdx = rest.lastIndexOf(" : ");
      if (colonIdx === -1) return;
      var label = rest.substring(0, colonIdx).trim();
      var valStr = rest.substring(colonIdx + 3).trim();
      var noteMatch = valStr.match(/^(\d+(?:\.\d+)?)\/10/);
      var note = noteMatch ? parseFloat(noteMatch[1]) : 0;
      if (valStr === "Oui") note = 1;
      answers.push({ qnum: qnum, label: label, value: valStr, note: note });
    }
  });
  saveToClients(ss, data, date);
  createPlanEntrainement(ss, data, answers, date, traumaSet);
}

function saveToClients(ss, data, date) {
  var sheet = ss.getSheetByName("Clients");
  if (!sheet) {
    sheet = ss.insertSheet("Clients");
    sheet.getRange(1,1,1,9).setValues([["Date","Nom","Prénom","Âge","DDN","Email","Téléphone","Ville","Raison"]])
      .setBackground("#3a2d8f").setFontColor("white").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([date, data.nom, data.prenom, data.age, data.ddn, data.email, data.tel, data.ville, data.raison]);
}

// 1=certain(X), 2=incertain(?)
// Colonnes: T3T4, T4P4, T4FP2, T3FP1, AUTRES, SYNCH, AT
var SYMPTOMES = [
  // ── SOMMEIL ──
  { qnum:2,  section:"Sommeil", label:"Sommeil non récupérateur",             T4P4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:3,  section:"Sommeil", label:"Sommeil agité — pb émotionnel",        T4P4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:3,  section:"Sommeil", label:"Sommeil agité — petit vélo (pensées)", T4P4:1, T3FP1:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:4,  section:"Sommeil", label:"Difficultés à s'endormir — agitation", T4P4:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:4,  section:"Sommeil", label:"Difficultés à s'endormir — hypervigilance/peurs", T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:4,  section:"Sommeil", label:"Difficultés à s'endormir — petit vélo (pensées)", T3FP1:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:5,  section:"Sommeil", label:"Difficultés à se réveiller — travail de nuit", T3T4:1, T4P4:1, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  { qnum:5,  section:"Sommeil", label:"Difficultés à se réveiller — dépression", T4P4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  { qnum:6,  section:"Sommeil", label:"Réveils fréquents — pb émotionnel",    T3T4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  { qnum:6,  section:"Sommeil", label:"Réveils fréquents — petit vélo",       T3T4:1, T3FP1:1, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  { qnum:7,  section:"Sommeil", label:"Sommeil irrégulier — pb émotionnel",   T4P4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:8,  section:"Sommeil", label:"Cauchemars / rêves vivants",           T4P4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:9,  section:"Sommeil", label:"Jambes agitées",                       T4P4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  { qnum:10, section:"Sommeil", label:"Jambes sans repos — stress/fatigue/addictions", T3T4:1, T4P4:1, AUTRES:2, SYNCH:1, AT:2 },
  { qnum:10, section:"Sommeil", label:"Jambes sans repos — pb émotionnel/addictions", T3T4:1, T4FP2:1, AUTRES:2, SYNCH:1, AT:2 },
  { qnum:11, section:"Sommeil", label:"Terreurs nocturnes",                   T3T4:1, T4FP2:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:12, section:"Sommeil", label:"Sueurs nocturnes",                     T3T4:1, T4FP2:2, AUTRES:2, SYNCH:1, AT:1 },
  { qnum:13, section:"Sommeil", label:"Apnées du sommeil",                    T3T4:1, T4FP2:2, T3FP1:2, SYNCH:2, AT:2 },
  { qnum:14, section:"Sommeil", label:"Ronflements",                          T4P4:1 },
  { qnum:15, section:"Sommeil", label:"Somniloquie",                          T3T4:1, T4FP2:2, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:16, section:"Sommeil", label:"Bruxisme",                             T4P4:1, T4FP2:2, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:17, section:"Sommeil", label:"Somnambulisme",                        T3T4:1, T4FP2:2, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:18, section:"Sommeil", label:"Énurésie",                             T4P4:1, T4FP2:2, AUTRES:"T4/O2 si trauma", SYNCH:2, AT:2 },
  // ── COGNITION ──
  { qnum:20, section:"Cognition", label:"Difficultés d'attention/concentration", T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"T3/P3 si attention aux détails", SYNCH:2, AT:2 },
  { qnum:21, section:"Cognition", label:"Difficultés calcul/maths/géométrie",    T4P4:1, AUTRES:"T3/P3", SYNCH:"40Hz", AT:2 },
  { qnum:22, section:"Cognition", label:"Écriture illisible",                    T4P4:1, AUTRES:"T3/P3" },
  { qnum:23, section:"Cognition", label:"Difficultés à dessiner",                T4P4:1, AUTRES:"T3/P3" },
  { qnum:24, section:"Cognition", label:"Difficultés à lire",                    T4P4:1, AUTRES:"T3/P3 + T3/F5" },
  { qnum:25, section:"Cognition", label:"Difficultés à planifier/s'organiser",   T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"40hz", SYNCH:2, AT:2 },
  { qnum:26, section:"Cognition", label:"Difficultés à terminer les tâches",     T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"40hz", SYNCH:2, AT:2 },
  { qnum:27, section:"Cognition", label:"Difficultés à suivre une conversation",  T4P4:1, T4FP2:2, AUTRES:"T3/F7 si expression verbale", AT:1 },
  { qnum:28, section:"Cognition", label:"Difficultés à comprendre/appliquer",    T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"T4/CP6 si compréhension du langage", SYNCH:2, AT:2 },
  { qnum:29, section:"Cognition", label:"Procrastination",                       T4P4:1, T4FP2:2, T3FP1:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:30, section:"Cognition", label:"Difficultés à prendre des décisions",   T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:31, section:"Cognition", label:"Difficultés à s'exprimer/articuler",    T4P4:1, T4FP2:2, T3FP1:1, AUTRES:"T3/F7 si expression verbale", SYNCH:1, AT:1 },
  { qnum:31, section:"Cognition", label:"Difficultés à s'exprimer — langage précoce",   T4P4:1, AUTRES:"T4/F8 si langage précoce", SYNCH:1, AT:1 },
  { qnum:32, section:"Cognition", label:"Difficultés à mémoriser les noms",      T4P4:1, T3FP1:1, AUTRES:"T3/P3", SYNCH:1, AT:2 },
  { qnum:33, section:"Cognition", label:"Difficultés mémoire court terme",       T4P4:1, AUTRES:"T3/P3", SYNCH:1 },
  { qnum:34, section:"Cognition", label:"Difficultés mémoire long terme",        T4P4:1, T3FP1:1, AUTRES:"T3/P3", SYNCH:1, AT:2 },
  { qnum:35, section:"Cognition", label:"Difficultés à assimiler de nouvelles connaissances", T4P4:1, T4FP2:2, AUTRES:"T3/P3", SYNCH:1, AT:2 },
  { qnum:36, section:"Cognition", label:"Démotivation dans l'apprentissage",     T4P4:1, T3FP1:1, AUTRES:"T3/P3", SYNCH:1, AT:2 },
  { qnum:37, section:"Cognition", label:"HPI diagnostiqué",                      T4P4:1, T3FP1:1, AUTRES:"T3/P3 + T3/F5", SYNCH:2, AT:2 },
  { qnum:38, section:"Cognition", label:"HPE diagnostiqué",                      T4P4:1, T4FP2:2, SYNCH:1, AT:2 },
  { qnum:39, section:"Cognition", label:"Troubles dys",                          T4P4:1, AUTRES:"T3/P3 Et T3/F5", SYNCH:2, AT:2 },
  // ── PERCEPTION SENSORIELLE ──
  { qnum:41, section:"Perception", label:"Peu attentif au corps/sensations",     T4P4:1, T4FP2:1, SYNCH:2, AT:2 },
  { qnum:42, section:"Perception", label:"Confusion gauche/droite",              T4P4:1, T3FP1:1 },
  { qnum:43, section:"Perception", label:"Conscience altérée de l'espace/temps", T4P4:1, T4FP2:2, AUTRES:1 },
  { qnum:44, section:"Perception", label:"Mauvais sens de l'orientation",        T4P4:1, T4FP2:2, AUTRES:1 },
  { qnum:45, section:"Perception", label:"Problèmes auditifs / hypersensibilité", T3T4:1, T4P4:1, T4FP2:2, SYNCH:2, AT:2 },
  { qnum:46, section:"Perception", label:"Acouphènes",                           T3T4:1, T4FP2:2, SYNCH:2, AT:2 },
  { qnum:47, section:"Perception", label:"Déficits visuels / hypersensibilité",  T4P4:1, T4FP2:2, AUTRES:"T4/O2 + O1/O2", SYNCH:2, AT:2 },
  { qnum:48, section:"Perception", label:"Hypersensibilité au toucher",          T4P4:1 },
  { qnum:49, section:"Perception", label:"Vertige / peur du vide",               T3T4:1 },
  { qnum:50, section:"Perception", label:"Intolérances chimiques",               T3T4:1 },
  { qnum:51, section:"Perception", label:"Mal des transports",                   T3T4:1 },
  { qnum:52, section:"Perception", label:"Fluctuations d'énergie",               T3T4:1, SYNCH:2, AT:2 },
  // ── COMPORTEMENTS ──
  { qnum:54, section:"Comportements", label:"Introverti",                        T4P4:1, T4FP2:1, SYNCH:"0.05mHz", AT:1 },
  { qnum:54, section:"Comportements", label:"Expansif / extraverti",             T4P4:1, T4FP2:1, SYNCH:"0.05mHz", AT:1 },
  { qnum:55, section:"Comportements", label:"Actif",                             T4P4:1, T4FP2:1, SYNCH:"0.05mHz", AT:1 },
  { qnum:55, section:"Comportements", label:"Rêveur",                            T4P4:1, T4FP2:1, SYNCH:"0.05mHz", AT:1 },
  { qnum:56, section:"Comportements", label:"Stressé / agité",                   T4P4:1, T4FP2:1, T3FP1:1, SYNCH:"0.05mHz 10Hz 40Hz", AT:1 },
  { qnum:56, section:"Comportements", label:"Anxieux",                           T4P4:1, T4FP2:1, SYNCH:"0.05mHz", AT:2 },
  { qnum:57, section:"Comportements", label:"Difficultés à se détendre",         T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:"0.05mHz 10Hz", AT:2 },
  { qnum:59, section:"Comportements", label:"Difficultés à s'affirmer",          T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:60, section:"Comportements", label:"Sens de l'humour",                  T4P4:1, T4FP2:1, SYNCH:1, AT:1 },
  { qnum:61, section:"Comportements", label:"Parle beaucoup",                    T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:63, section:"Comportements", label:"Impulsivité",                       T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:64, section:"Comportements", label:"Comportement compulsif",            T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:65, section:"Comportements", label:"Opposition aux autres",             T4P4:1, T4FP2:1, T3FP1:2, AUTRES:"T4/O2", SYNCH:1, AT:2 },
  { qnum:66, section:"Comportements", label:"Agressivité",                       T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:67, section:"Comportements", label:"Crises de colère / rage",           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:68, section:"Comportements", label:"Auto-mutilation",                   T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:69, section:"Comportements", label:"Addiction aliments sucrés/café",    T3T4:1, T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:2 },
  { qnum:69, section:"Comportements", label:"Addiction alcool/tabac/régimes",    T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:2 },
  { qnum:69, section:"Comportements", label:"Addiction marijuana/drogues",       T3T4:1, T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:2 },
  { qnum:69, section:"Comportements", label:"Addiction écrans/autres",           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:2 },
  { qnum:70, section:"Comportements", label:"TCA (anorexie/boulimie)",           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:71, section:"Comportements", label:"Manies",                            T4P4:1, T3FP1:1, SYNCH:2, AT:2 },
  { qnum:72, section:"Comportements", label:"Tics physiques ou vocaux",          T4P4:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:73, section:"Comportements", label:"TOC",                               T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:74, section:"Comportements", label:"TDAH diagnostiqué",                 T3T4:1, T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:75, section:"Comportements", label:"Comportements à risques",           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:76, section:"Comportements", label:"Manipulation",                      T4P4:1, T4FP2:1, T3FP1:1, SYNCH:1, AT:1 },
  { qnum:77, section:"Comportements", label:"Anxiété sociale",                   T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:78, section:"Comportements", label:"Désorganisé",                       T4P4:1, T4FP2:2, T3FP1:1, SYNCH:2, AT:2 },
  { qnum:79, section:"Comportements", label:"Attention à l'hygiène",               T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:80, section:"Comportements", label:"Méfiance généralisée",              T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:81, section:"Comportements", label:"Autisme — expression émotionnelle",   T3T4:1, T4FP2:1, AUTRES:"T4/F8 si expression émotionnelle", SYNCH:1, AT:1 },
  { qnum:81, section:"Comportements", label:"Autisme — expression faciale/langage",    T3T4:1, AUTRES:"T4/T6 si expression faciale", SYNCH:1, AT:1 },
  { qnum:81, section:"Comportements", label:"Autisme — si trauma",                      T3T4:1, AUTRES:"T4/O2 si trauma", SYNCH:1, AT:1 },
  { qnum:82, section:"Comportements", label:"TSPT",                              T3T4:1, T4P4:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:83, section:"Comportements", label:"Trouble bipolaire",                 T3T4:1, T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:84, section:"Comportements", label:"Difficultés sexualité",             T3T4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  // ── EMOTIONNEL ──
  { qnum:87, section:"Emotionnel", label:"Manque de confiance en soi",           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:88, section:"Emotionnel", label:"Faible estime de soi",                 T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:89, section:"Emotionnel", label:"Craintes / appréhensions",             T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:90, section:"Emotionnel", label:"Peurs",                                T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:91, section:"Emotionnel", label:"Phobies",                              T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:92, section:"Emotionnel", label:"Embarras / gêne en public",            T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:93, section:"Emotionnel", label:"Méfiance vis-à-vis des autres",        T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:94, section:"Emotionnel", label:"Attaques de panique",                  T3T4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:96, section:"Emotionnel", label:"Manque de joie / plaisirs",            T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:95, section:"Emotionnel", label:"Tristesse / dépression",               T4P4:1, T4FP2:1, T3FP1:2, AUTRES:"T4/O2 T3/F3 effet anti-dépresseur", SYNCH:1, AT:1 },
  { qnum:97, section:"Emotionnel", label:"Empathie (peu ou trop)",               T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:98, section:"Emotionnel", label:"Sautes d'humeur",                      T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:99, section:"Emotionnel", label:"Impatience",                           T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:100,section:"Emotionnel", label:"Hypervigilance émotionnelle",          T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:101,section:"Emotionnel", label:"Irritabilité quand contrarié",         T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:102,section:"Emotionnel", label:"Difficultés à se calmer",              T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:103,section:"Emotionnel", label:"Pensées obsessionnelles",              T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:104,section:"Emotionnel", label:"Pensées négatives obsessionnelles",    T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:105,section:"Emotionnel", label:"Souvenirs de traumatismes",            T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:106,section:"Emotionnel", label:"Dépression diagnostiquée / burn out",  T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:107,section:"Emotionnel", label:"Burn out",                             T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:108,section:"Emotionnel", label:"Déconnexion de la réalité",            T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:109,section:"Emotionnel", label:"Dissociation",                         T3T4:1, T4FP2:1, AUTRES:"T4/O2 T3/F3", SYNCH:1, AT:1 },
  { qnum:110,section:"Emotionnel", label:"Hallucinations",                       T4P4:1, T3FP1:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:111,section:"Emotionnel", label:"Pensées suicidaires",                  T4P4:1, T4FP2:1 },
  // ── PHYSIQUE ──
  { qnum:113,section:"Physique", label:"Fatigue chronique",                      T3T4:1, T4FP2:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:113,section:"Physique", label:"Fatigue rapide après effort",            T4P4:2, T4FP2:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:113,section:"Physique", label:"Épuisé(e) en ce moment",                T4P4:2 },
  { qnum:114,section:"Physique", label:"Tensions musculaires",                   T4P4:1, T4FP2:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:115,section:"Physique", label:"Faiblesse motrice — si AVC",            T3T4:1, T4P4:1, AUTRES:"C3/C4 si AVC" },
  { qnum:116,section:"Physique", label:"Manque de coordination — si AVC",       T3T4:1, T4P4:1, AUTRES:"C3/C4 T3/P3 si AVC", SYNCH:2, AT:2 },
  { qnum:117,section:"Physique", label:"Tics musculaires",                       T4P4:1, T3FP1:1 },
  { qnum:118,section:"Physique", label:"Crampes musculaires",                    T4P4:1, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:119,section:"Physique", label:"Difficultés physiques au travail",       T4P4:1 },
  { qnum:120,section:"Physique", label:"Spasticité",                             T4P4:1 },
  { qnum:121,section:"Physique", label:"Troubles de l'équilibre",                T4P4:1 },
  { qnum:122,section:"Physique", label:"Tremblements",                           T4P4:1 },
  { qnum:123,section:"Physique", label:"Maladie de Parkinson",                   T3T4:1, T4P4:1 },
  { qnum:124,section:"Physique", label:"Allergies",                              T3T4:1 },
  { qnum:125,section:"Physique", label:"Asthme",                                 T3T4:1 },
  { qnum:126,section:"Physique", label:"Intestin irritable",                     T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:127,section:"Physique", label:"Constipation chronique",                 T3T4:1 },
  { qnum:128,section:"Physique", label:"Diarrhées fréquentes",                   T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:129,section:"Physique", label:"Reflux / brûlures d'estomac",            T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:130,section:"Physique", label:"Digestion aliments sucrés difficile",    T3T4:1 },
  { qnum:131,section:"Physique", label:"Nausées",                                T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:132,section:"Physique", label:"Vertiges / évanouissements",             T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:133,section:"Physique", label:"Peau irritée / réactive",                T3T4:1 },
  { qnum:134,section:"Physique", label:"SPM — bouffées de chaleur/migraines",    T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:134,section:"Physique", label:"SPM — si agressivité",                   T4P4:1, T4FP2:1, T3FP1:1, AUTRES:"X si brouillard cérébral", SYNCH:2, AT:2 },
  { qnum:135,section:"Physique", label:"Bouffées de chaleur",                    T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:136,section:"Physique", label:"Ménopause / andropause",                 T3T4:1 },
  { qnum:137,section:"Physique", label:"Pb liés à la ménopause/andropause",      T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:138,section:"Physique", label:"Transpire en cas de stress",              T3T4:1 },
  { qnum:139,section:"Physique", label:"Transpire fréquemment",                  T3T4:1 },
  { qnum:140,section:"Physique", label:"Incontinences — émotionnel",             T3T4:1, T4FP2:1 },
  
  { qnum:141,section:"Physique", label:"Incontinences — pb physique",            T3T4:1, T4P4:1 },
  { qnum:142,section:"Physique", label:"Arythmie / tachycardie — pb physique",   T4P4:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:142,section:"Physique", label:"Arythmie / tachycardie — si émotionnel", T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:143,section:"Physique", label:"Hypertension / hypotension — si émot.",  T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:144,section:"Physique", label:"Problème de thyroïde — si émot.",        T4P4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:145,section:"Physique", label:"Dérégulations endocriniennes",           T3T4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:146,section:"Physique", label:"Diabète — si émotionnel",                T3T4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:147,section:"Physique", label:"Narcolepsie",                            T3T4:1, T4FP2:2, SYNCH:2, AT:2 },
  { qnum:148,section:"Physique", label:"Maux de tête récurrents",                T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:149,section:"Physique", label:"Maux de tête — tensions musculaires",     T4P4:1 },
  { qnum:150,section:"Physique", label:"Migraines",                              T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:151,section:"Physique", label:"Névralgies",                             T4P4:1 },
  { qnum:152,section:"Physique", label:"Douleurs neuropathiques",                T4P4:1 },
  { qnum:153,section:"Physique", label:"Névralgies du trijumeau",                T3T4:1 },
  { qnum:154,section:"Physique", label:"Névralgies d'Arnold",                    T3T4:1 },
  { qnum:155,section:"Physique", label:"Sciatique",                              T4P4:1 },
  { qnum:156,section:"Physique", label:"Douleurs des sinus",                     T3T4:1 },
  { qnum:157,section:"Physique", label:"Douleurs musculaires",                   T4P4:1 },
  { qnum:158,section:"Physique", label:"Fibromyalgie",                           T3T4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:159,section:"Physique", label:"Douleurs d'estomac",                     T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:160,section:"Physique", label:"Douleurs intestinales",                  T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:161,section:"Physique", label:"Douleurs urinaires",                     T3T4:2, T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:162,section:"Physique", label:"Tolérance à la douleur (faible)",         T4P4:1 },
  { qnum:163,section:"Physique", label:"Maladies en ite (méningite, gastrite…)", T3T4:1, T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:164,section:"Physique", label:"Épilepsie / pseudo-crises",              T3T4:1, T4P4:1, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:165,section:"Physique", label:"Faiblesse immunitaire",                  T3T4:1, T4P4:1, AUTRES:"T3/P3", SYNCH:2, AT:2 },
  { qnum:166,section:"Physique", label:"Covid long",                             T3T4:1, T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:2, AT:2 },
  { qnum:167,section:"Physique", label:"Maladie de Charcot",                     T3T4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:168,section:"Physique", label:"Sclérose en plaques",                    T3T4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:169,section:"Physique", label:"VIH",                                    T3T4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:170,section:"Physique", label:"Cancer",                                 T4P4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:171,section:"Physique", label:"Tumeur au cerveau",                      T3T4:1, T4FP2:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:172,section:"Physique", label:"Encéphalites",                           T3T4:1, T4FP2:2, T4P4:2, T3FP1:2, AUTRES:2, SYNCH:2, AT:2 },
  // ── ANTECEDENTS ──
  { qnum:173,section:"Physique", label:"Traumatisme crânien",                 T3T4:1, T4P4:2, T4FP2:2, T3FP1:1, AUTRES:2, SYNCH:2, AT:2 },
  { qnum:174,section:"Physique", label:"AVC — hémisphère gauche",             T3T4:1, T4P4:2, T4FP2:2, SYNCH:2, AT:2 },
  { qnum:174,section:"Physique", label:"AVC — hémisphère droit",              T3T4:1, T4P4:2, T4FP2:2, SYNCH:2, AT:2 },
  { qnum:175,section:"Physique", label:"Autres problèmes cérébraux",          T3T4:1, T4P4:2, T3FP1:1, SYNCH:2, AT:2 },
  { qnum:176,section:"Petite enfance", label:"Traumatisme in utero / naissance",    T3T4:1, T4FP2:1, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:177,section:"Petite enfance", label:"Pb développement physique/moteur",    T4P4:1, T4FP2:2, AUTRES:"???", SYNCH:2, AT:2 },
  { qnum:178,section:"Petite enfance", label:"Convulsions / épilepsie",             T3T4:1, T4P4:1, T4FP2:2, SYNCH:1, AT:1 },
  { qnum:179,section:"Petite enfance", label:"Traumatismes précoces (abus/violence)", T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:180,section:"Petite enfance", label:"Adopté(e)",                           T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:181,section:"Petite enfance", label:"Pb d'attachement / relations difficiles", T4P4:1, T4FP2:2, AUTRES:"T4/O2", SYNCH:1, AT:1 },
  { qnum:182,section:"Petite enfance", label:"Pb naissance ou d'enfance",           T4P4:1, AUTRES:"T4/O2", SYNCH:1, AT:1 }
];

function createPlanEntrainement(ss, data, answers, date, traumaSet) {
  traumaSet = traumaSet || {};
  var nom = (data.prenom + " " + data.nom).trim();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var sheetName = (nom || "Client").substring(0, 25) + " " + dateStr;
  var sheet;
  try { sheet = ss.insertSheet(sheetName); }
  catch(e) { sheet = ss.insertSheet(sheetName + " " + Math.floor(Math.random()*99)); }

  var V="#3a2d8f", VL="#eae5f7", W="#ffffff", OR="#ffe0b2", RD="#ffcccc", GR="#f5f5f5";
  var r=1;

  sheet.getRange(r,1,1,4).merge().setValue("PLAN D ENTRAINEMENT DES SEANCES DE NEUROFEEDBACK en ILF")
    .setBackground(V).setFontColor(W).setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");
  r++;
  sheet.getRange(r,1,1,4).setBackground(VL);
  sheet.getRange(r,1).setValue("Nom :").setFontWeight("bold");
  sheet.getRange(r,2).setValue(nom).setFontWeight("bold").setFontSize(13);
  sheet.getRange(r,3).setValue("Date :").setFontWeight("bold");
  sheet.getRange(r,4).setValue(dateStr);
  r++;
  sheet.getRange(r,1).setValue("Age :").setFontWeight("bold");
  sheet.getRange(r,2).setValue(data.age);
  sheet.getRange(r,3).setValue("Email :").setFontWeight("bold");
  sheet.getRange(r,4).setValue(data.email);
  r++;
  sheet.getRange(r,1).setValue("Raison :").setFontWeight("bold");
  sheet.getRange(r,2,1,3).merge().setValue(data.raison).setWrap(true);
  r++; r++;

  var answerMap = {};
  answers.forEach(function(a) { answerMap[a.qnum] = a; });

  var caps = { T3FP1:[], T3T4:[], T4FP2:[], T3P3:[], T4P4:[], AUTRES:[], SYNCH:[], AT:[] };
  SYMPTOMES.forEach(function(s) {
    var a = answerMap[s.qnum];
    if (!a || a.note === 0) return;
    var e = { label:s.label, note:a.note, section:s.section };
    if (s.T3FP1) caps.T3FP1.push(Object.assign({},e,{uncertain:s.T3FP1===2}));
    if (s.T3T4)  caps.T3T4.push(Object.assign({},e,{uncertain:s.T3T4===2}));
    if (s.T4FP2) caps.T4FP2.push(Object.assign({},e,{uncertain:s.T4FP2===2}));
    if (s.T3P3)  caps.T3P3.push(Object.assign({},e,{uncertain:s.T3P3===2}));
    if (s.T4P4)  caps.T4P4.push(Object.assign({},e,{uncertain:s.T4P4===2}));
    // AUTRES: conditionnel si contient "si "
    var autresVal = s.AUTRES;
    if (typeof autresVal === "string" && autresVal.toLowerCase().indexOf("si ") !== -1) {
      // Extraire le label de la condition ex: "T4/O2 si trauma" -> condition = "trauma", capteur = "T4/O2"
      // "T3/P3 si attention aux détails" -> condition = "attention aux détails", capteur = "T3/P3"
      var siIdx = autresVal.toLowerCase().indexOf(" si ");
      var capteur = autresVal.substring(0, siIdx).trim();
      var condition = autresVal.substring(siIdx + 4).trim().toLowerCase();
      // Vérifier si le praticien a coché cette condition pour ce qnum
      var conditionKey = s.qnum + "_" + condition;
      if (traumaSet[conditionKey] || (condition === "trauma" && traumaSet[s.qnum])) {
        caps.AUTRES.push(Object.assign({},e,{uncertain:false, note2:capteur}));
      }
    } else if (autresVal && autresVal !== 0) {
      caps.AUTRES.push(Object.assign({},e,{uncertain:autresVal===2, note2:typeof autresVal==="string"?autresVal:""}));
    }
    if (s.SYNCH && s.SYNCH !== 0) caps.SYNCH.push(Object.assign({},e,{uncertain:s.SYNCH===2, note2:typeof s.SYNCH==="string"?s.SYNCH:""}));
    if (s.AT && s.AT !== 0) caps.AT.push(Object.assign({},e,{uncertain:s.AT===2}));
  });

  function writeItems(col, items) {
    var added = 0;
    if (items.length === 0) {
      sheet.getRange(r, col, 2, 2).merge().setBackground(W); return 2;
    }
    var curSec = "";
    items.forEach(function(item) {
      if (item.section !== curSec) {
        curSec = item.section;
        sheet.getRange(r+added, col, 1, 2).merge()
          .setValue(curSec).setBackground("#d8d0f0").setFontColor(V)
          .setFontWeight("bold").setFontSize(8).setHorizontalAlignment("center");
        added++;
      }
      var rowBg = item.uncertain ? "#e0e0e0" : W;
      var lbl = item.uncertain ? item.label + " (?)" : item.label;
      if (item.note2) lbl = lbl + " [" + item.note2 + "]";
      var lc = sheet.getRange(r+added, col);
      lc.setValue(lbl).setFontSize(9).setWrap(true).setBackground(rowBg);
      if (item.uncertain) lc.setFontColor("#999999");
      var nc = sheet.getRange(r+added, col+1);
      if (item.uncertain) {
        nc.setValue(item.note).setHorizontalAlignment("center").setFontSize(11).setBackground(rowBg).setFontColor("#999999");
      } else {
        nc.setValue(item.note).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
        if (item.note >= 7) { nc.setBackground(RD); nc.setFontColor("#c0392b"); }
        else if (item.note >= 5) { nc.setBackground(OR); nc.setFontColor("#e67e22"); }
      }
      added++;
    });
    return added;
  }

  function writeBlock(col, title, subtitle, items, color) {
    sheet.getRange(r, col, 1, 2).merge().setValue(title)
      .setBackground(color).setFontColor(W).setFontWeight("bold").setFontSize(10);
    r++;
    if (subtitle) {
      sheet.getRange(r, col, 1, 2).merge().setValue(subtitle)
        .setBackground(VL).setFontColor(V).setFontStyle("italic").setFontSize(8).setWrap(true);
      r++;
    }
    var added = writeItems(col, items);
    return r + added;
  }

  sheet.getRange(r,1,1,4).merge().setValue("1)  EMPLACEMENTS PRINCIPAUX")
    .setBackground(V).setFontColor(W).setFontSize(11).setFontWeight("bold");
  r++; r++;

  var savedR = r;
  var endL = writeBlock(1, "T3/FP1  Avant Gauche", "Calmer le mental et le controle des impulsions", caps.T3FP1, "#1a237e");
  var tempR = r;
  r = savedR;
  var endR = writeBlock(3, "T3/T4  Stabiliser", null, caps.T3T4, "#283593");
  r = Math.max(endL, endR); r++;

  savedR = r;
  endL = writeBlock(1, "T4/FP2  Avant Droit", "Calmer l emotionnel", caps.T4FP2, "#1565c0");
  tempR = r;
  r = savedR;
  endR = writeBlock(3, "T4/P4  Arriere Droit", "Calmer le physique et la conscience corporelle", caps.T4P4, "#00695c");
  r = Math.max(endL, endR); r++;

  sheet.getRange(r,1,1,4).merge().setValue("T3/P3  Arriere Gauche  -  Orientation vers le detail, les competences et la connaissance")
    .setBackground("#0277bd").setFontColor(W).setFontWeight("bold").setFontSize(10);
  r++;
  if (caps.T3P3.length === 0) {
    sheet.getRange(r,1,2,4).merge().setBackground(W); r+=2;
  } else {
    caps.T3P3.forEach(function(item) {
      sheet.getRange(r,1,1,3).merge().setValue(item.label).setFontSize(9).setWrap(true);
      var nc = sheet.getRange(r,4);
      nc.setValue(item.note).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
      if (item.note >= 7) { nc.setBackground(RD); nc.setFontColor("#c0392b"); }
      else if (item.note >= 5) { nc.setBackground(OR); nc.setFontColor("#e67e22"); }
      r++;
    });
  }
  r++;

  sheet.getRange(r,1,1,4).merge().setValue("2)  AUTRES EMPLACEMENTS")
    .setBackground(V).setFontColor(W).setFontSize(11).setFontWeight("bold");
  r++;
  if (caps.AUTRES.length === 0) { sheet.getRange(r,1,3,4).merge().setBackground(W); r+=3; }
  else {
    caps.AUTRES.forEach(function(item) {
      var rowBg = item.uncertain ? "#e0e0e0" : W;
      var lbl = item.note2 ? item.label + " [" + item.note2 + "]" : item.label;
      sheet.getRange(r,1,1,3).merge().setValue(lbl).setFontSize(9).setWrap(true).setBackground(rowBg);
      var nc = sheet.getRange(r,4);
      nc.setValue(item.note).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
      if (!item.uncertain) {
        if (item.note >= 7) { nc.setBackground(RD); nc.setFontColor("#c0392b"); }
        else if (item.note >= 5) { nc.setBackground(OR); nc.setFontColor("#e67e22"); }
      } else { nc.setBackground(rowBg).setFontColor("#999999"); }
      r++;
    });
  }
  r++;

  sheet.getRange(r,1,1,4).merge().setValue("3)  SYNCHRONIE ET ALPHA/THETA")
    .setBackground(V).setFontColor(W).setFontSize(11).setFontWeight("bold");
  r++;
  sheet.getRange(r,1,1,2).merge().setValue("Synchronie").setFontWeight("bold").setBackground(VL).setFontColor(V);
  sheet.getRange(r,3,1,2).merge().setValue("Alpha/Theta").setFontWeight("bold").setBackground(VL).setFontColor(V);
  r++;
  var maxR = Math.max(caps.SYNCH.length, caps.AT.length, 3);
  for (var i=0; i<maxR; i++) {
    var bg = i%2===0 ? GR : W;
    var s = caps.SYNCH[i]; var a = caps.AT[i];
    sheet.getRange(r,1).setValue(s ? s.label : "").setBackground(bg).setFontSize(9).setWrap(true);
    if (s && s.note > 0) {
      var sc = sheet.getRange(r,2);
      sc.setValue(s.note).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
      if (!s.uncertain) {
        if (s.note >= 7) { sc.setBackground(RD); sc.setFontColor("#c0392b"); }
        else if (s.note >= 5) { sc.setBackground(OR); sc.setFontColor("#e67e22"); }
      } else { sc.setBackground("#e0e0e0").setFontColor("#999999"); }
    } else { sheet.getRange(r,2).setBackground(bg); }
    sheet.getRange(r,3).setValue(a ? a.label : "").setBackground(bg).setFontSize(9).setWrap(true);
    if (a && a.note > 0) {
      var ac = sheet.getRange(r,4);
      ac.setValue(a.note).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
      if (!a.uncertain) {
        if (a.note >= 7) { ac.setBackground(RD); ac.setFontColor("#c0392b"); }
        else if (a.note >= 5) { ac.setBackground(OR); ac.setFontColor("#e67e22"); }
      } else { ac.setBackground("#e0e0e0").setFontColor("#999999"); }
    } else { sheet.getRange(r,4).setBackground(bg); }
    r++;
  } 
 
  // ── Réaction au stress (Q58) et Réaction à la détente (Q58b) ── 
  r++;
  sheet.getRange(r,1,1,4).merge().setValue("4)  RÉACTION AU STRESS ET À LA DÉTENTE") 
    .setBackground(V).setFontColor(W).setFontSize(11).setFontWeight("bold"); 
  r++;
  var q58 = answerMap["58"] || answerMap[58]; 
  var q58b = answerMap["58b"]; 
  sheet.getRange(r,1,1,4).merge().setValue("Réaction au stress") 
    .setBackground(VL).setFontColor(V).setFontWeight("bold").setFontSize(10); 
  r++;
  var stress_text = (q58 && q58.comment) ? q58.comment : (q58 && q58.text ? q58.text : "—"); 
  sheet.getRange(r,1,1,4).merge().setValue(stress_text) 
    .setFontSize(9).setWrap(true).setBackground(W); 
  r++; r++;
  sheet.getRange(r,1,1,4).merge().setValue("Réaction à la détente") 
    .setBackground(VL).setFontColor(V).setFontWeight("bold").setFontSize(10); 
  r++;
  var detente_text = (q58b && q58b.comment) ? q58b.comment : (q58b && q58b.text ? q58b.text : "—"); 
  sheet.getRange(r,1,1,4).merge().setValue(detente_text) 
    .setFontSize(9).setWrap(true).setBackground(W); 
  r++;
 
  sheet.setColumnWidth(1,260); sheet.setColumnWidth(2,55); 
  sheet.setColumnWidth(3,260); sheet.setColumnWidth(4,55); 
  sheet.setFrozenRows(5); 
}

function createTableauReference() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName("Tableau de reference");
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet("Tableau de reference");
  var headers = ["N", "Categorie", "Symptome", "T3/T4", "T4/P4", "T3/FP1", "T4/FP2", "AUTRES", "SYNCH", "A/T"];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground("#3a2d8f").setFontColor("white").setFontWeight("bold");
  sheet.setFrozenRows(1);
  var r=2; var curSec="";
  SYMPTOMES.forEach(function(s) {
    if (s.section !== curSec) {
      curSec = s.section;
      sheet.getRange(r,1,1,10).merge().setValue(curSec.toUpperCase())
        .setBackground("#eae5f7").setFontColor("#3a2d8f").setFontWeight("bold").setHorizontalAlignment("center");
      r++;
    }
    var bg = r%2===0 ? "#f8f6ff" : "#ffffff";
    sheet.getRange(r,1).setValue(s.qnum).setHorizontalAlignment("center").setBackground(bg);
    sheet.getRange(r,2).setValue(s.section).setBackground(bg).setFontColor("#666");
    sheet.getRange(r,3).setValue(s.label).setBackground(bg);
    // Order: T3T4, T4P4, T3FP1, T4FP2, AUTRES, SYNCH, AT
    var vals = [s.T3T4, s.T4P4, s.T3FP1, s.T4FP2, s.AUTRES, s.SYNCH, s.AT];
    vals.forEach(function(v,i) {
      var cell = sheet.getRange(r,4+i);
      cell.setBackground(bg);
      if (v === 1) cell.setValue("X").setHorizontalAlignment("center").setFontColor("#c0392b").setFontWeight("bold");
      else if (v === 2) cell.setValue("?").setHorizontalAlignment("center").setFontColor("#999999").setFontStyle("italic");
      else if (typeof v === "string" && v.length > 0) cell.setValue(v).setFontSize(7).setFontColor("#0277bd").setWrap(true);
    });
    r++;
  });
  sheet.setColumnWidth(1,40); sheet.setColumnWidth(2,100); sheet.setColumnWidth(3,260);
  for (var i=4; i<=10; i++) sheet.setColumnWidth(i,80);
  Logger.log("Tableau de reference cree !");
}

function setupSpreadsheet() {
  Logger.log("Configuration terminee !");
}
