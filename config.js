// Ρυθμίσεις Praxis CRM — ξεχωριστή εφαρμογή Entra και ξεχωριστές λίστες SharePoint (όχι του Tickets).
window.CRM_CONFIG = {
  tenantId: "83452b18-517e-4b31-a9e8-3cd60d2a8364",
  clientId: "76ec1e3c-688e-46f5-9086-3bad83f74858",   // Application (client) ID της εφαρμογής «Praxis CRM» στο Entra
  siteUrl:  "https://reinaldogr.sharepoint.com/sites/praxis",
  prospectsList: "CRM_Prospects",
  callsList: "CRM_Calls",
  auditsList: "CRM_Audits"   // έλεγχοι IT (δημιουργείται αυτόματα στην πρώτη αποθήκευση)
};
