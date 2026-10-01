// Avertit à l'envoi si le texte écrit annonce une pièce jointe absente.

const MOTS_CLES = [
  /\bci[- ]?joint(e|s|es)?\b/,
  /\bp\.? ?j\.?\b/,
  /\bpieces? jointes?\b/,
  /\ben annexe\b/,
  /\bannexe(e|s|es)?\b/,
  /\bsous ce pli\b/,
  /\bje vous (adresse|transmets|communique|fais parvenir)\b/,
  /\bvous trouverez\b/,
  /\b(veuillez|merci de) trouver\b/,
  /\bbordereau( de pieces)?\b/,
  /\ben attachement\b/,
  /\bjoint au present\b/,
];

// Lignes qui marquent le début du message cité (réponse / transfert).
const DEBUT_CITATION = [
  /^-{2,}\s*(message d'origine|original message|message transfere|forwarded message)/,
  /^_{5,}/,
  /^(de|from)\s*:/,
  /^le .{5,200} a ecrit\s*:?$/,
  /^on .{5,200} wrote\s*:?$/,
  /^>/,
];

function normaliser(texte) {
  return texte.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

function texteEcrit(corps) {
  const lignes = normaliser(corps).split(/\r?\n/);
  const garde = [];
  for (const ligne of lignes) {
    if (DEBUT_CITATION.some((re) => re.test(ligne.trim()))) break;
    garde.push(ligne);
  }
  return garde.join("\n");
}

const EXT_IMAGE = /\.(png|jpe?g|gif|bmp|svg|webp|emz|wmz)$/i;

// Une pièce compte si elle n'est pas une petite image intégrée (logo de signature…).
// Outlook Mac marque souvent les PDF ou les fichiers glissés comme « inline » : on les compte.
function estVraiePJ(a) {
  if (!a.isInline) return true;
  if (!EXT_IMAGE.test(a.name || "")) return true;
  return (a.size || 0) > 100000;
}

function annoncePJ(sujet, corps) {
  const texte = normaliser(sujet || "") + "\n" + texteEcrit(corps || "");
  return MOTS_CLES.some((re) => re.test(texte));
}

function doitAlerter(sujet, corps, pjs) {
  return annoncePJ(sujet, corps) && !(pjs || []).some(estVraiePJ);
}

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  item.subject.getAsync((resSujet) => {
    const sujet = resSujet.status === Office.AsyncResultStatus.Succeeded ? resSujet.value : "";
    item.body.getAsync(Office.CoercionType.Text, (resCorps) => {
      if (resCorps.status !== Office.AsyncResultStatus.Succeeded || !annoncePJ(sujet, resCorps.value)) {
        event.completed({ allowEvent: true });
        return;
      }
      item.getAttachmentsAsync((resPJ) => {
        if (resPJ.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({ allowEvent: true });
          return;
        }
        if (!doitAlerter(sujet, resCorps.value, resPJ.value)) {
          event.completed({ allowEvent: true });
          return;
        }
        event.completed({
          allowEvent: false,
          errorMessage:
            "Votre mail semble annoncer une pièce jointe (« ci-joint », « PJ », « en annexe »…), mais aucune n'est jointe." +
            " [diag : " + resPJ.value.length + " élément(s) vu(s)]",
        });
      });
    });
  });
}

if (typeof Office !== "undefined") {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
if (typeof module !== "undefined") {
  module.exports = { annoncePJ, doitAlerter, estVraiePJ, texteEcrit };
}
