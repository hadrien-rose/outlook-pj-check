const { doitAlerter } = require("./launchevent.js");
const pdf = { name: "contrat.pdf", isInline: false, size: 200000 };
const pdfInlineMac = { name: "contrat.pdf", isInline: true, size: 200000 };
const logo = { name: "image001.png", isInline: true, size: 8000 };
const photoInline = { name: "IMG_0042.jpg", isInline: true, size: 1500000 };
const cas = [
  // [description, sujet, corps, pièces, alerte attendue]
  ["ci-joint sans PJ", "", "Bonjour, ci-joint le contrat.", [], true],
  ["ci-joint avec 1 PJ", "", "Bonjour, ci-joint le contrat.", [pdf], false],
  ["ci-joint avec 2 PJ", "", "Ci-joints les deux actes.", [pdf, pdf], false],
  ["PDF marqué inline (Mac)", "", "Ci-joint le PDF.", [pdfInlineMac], false],
  ["2 PDF inline (Mac)", "", "Vous trouverez les pièces.", [pdfInlineMac, pdfInlineMac], false],
  ["seulement logo signature", "", "Ci-joint la note.", [logo], true],
  ["photo inline lourde", "", "Ci-joint la photo.", [photoInline], false],
  ["PJ dans l'objet", "Contrat - PJ", "Bonjour", [], true],
  ["P.J. avec points", "", "Voir P.J.", [], true],
  ["pièces jointes accent", "", "Les pièces jointes suivent.", [], true],
  ["en annexe", "", "Le devis est en annexe.", [], true],
  ["sous ce pli", "", "Sous ce pli, l'assignation.", [], true],
  ["je vous adresse", "", "Je vous adresse mes conclusions.", [], true],
  ["je vous transmets", "", "Je vous transmets le jugement.", [], true],
  ["veuillez trouver", "", "Veuillez trouver la facture.", [], true],
  ["bordereau", "", "Mon bordereau de pièces.", [], true],
  ["majuscules", "", "CI-JOINT LE DOSSIER", [], true],
  ["aucun mot-clé", "", "Bonne journée, à demain.", [], false],
  ["mot cité dans réponse", "", "Merci !\n\nLe 3 oct. 2026 à 10:00, Me X a écrit :\nci-joint", [], false],
  ["mot cité après De :", "", "Bien reçu.\n\nDe : X\nEnvoyé : …\nCi-joint", [], false],
  ["mot cité après >", "", "Ok\n> ci-joint", [], false],
  ["mot écrit + citation", "", "Ci-joint ma réponse.\n\nDe : X\nbla", [], true],
];
let ko = 0;
for (const [desc, sujet, corps, pjs, attendu] of cas) {
  const r = doitAlerter(sujet, corps, pjs);
  if (r !== attendu) ko++;
  console.log((r === attendu ? "OK  " : "ÉCHEC") + "  " + desc + "  → alerte=" + r);
}
console.log(`\n${cas.length - ko}/${cas.length} réussis`);
process.exit(ko ? 1 : 0);
