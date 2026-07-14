// Friendly product names for common Microsoft license SKUs (skuPartNumber).
// Not exhaustive — unknown SKUs fall back to a prettified part number.
const SKU_NAMES: Record<string, string> = {
  O365_BUSINESS_ESSENTIALS: 'Microsoft 365 Business Basic',
  O365_BUSINESS_PREMIUM: 'Microsoft 365 Business Standard',
  SPB: 'Microsoft 365 Business Premium',
  O365_BUSINESS: 'Microsoft 365 Apps for Business',
  OFFICESUBSCRIPTION: 'Microsoft 365 Apps for Enterprise',
  ENTERPRISEPACK: 'Office 365 E3',
  ENTERPRISEPREMIUM: 'Office 365 E5',
  STANDARDPACK: 'Office 365 E1',
  SPE_E3: 'Microsoft 365 E3',
  SPE_E5: 'Microsoft 365 E5',
  SPE_F1: 'Microsoft 365 F3',
  DESKLESSPACK: 'Office 365 F3',
  EXCHANGESTANDARD: 'Exchange Online (Plan 1)',
  EXCHANGEENTERPRISE: 'Exchange Online (Plan 2)',
  EXCHANGEDESKLESS: 'Exchange Online Kiosk',
  MCOSTANDARD: 'Skype for Business Online (Plan 2)',
  MCOEV: 'Microsoft Teams Phone Standard',
  MCOMEETADV: 'Microsoft 365 Audio Conferencing',
  PHONESYSTEM_VIRTUALUSER: 'Teams Phone Resource Account',
  TEAMS_EXPLORATORY: 'Microsoft Teams Exploratory',
  Microsoft_Teams_Premium: 'Microsoft Teams Premium',
  POWER_BI_STANDARD: 'Power BI (free)',
  POWER_BI_PRO: 'Power BI Pro',
  PBI_PREMIUM_PER_USER: 'Power BI Premium Per User',
  FLOW_FREE: 'Power Automate (free)',
  POWERAUTOMATE_ATTENDED_RPA: 'Power Automate Premium',
  POWERAPPS_VIRAL: 'Power Apps Plan 2 Trial',
  POWERAPPS_PER_USER: 'Power Apps Premium',
  CCIBOTS_PRIVPREV_VIRAL: 'Power Virtual Agents Trial',
  Power_Pages_vTrial_for_Makers: 'Power Pages vTrial for Makers',
  DYN365_ENTERPRISE_SALES: 'Dynamics 365 Sales Enterprise',
  PROJECTPROFESSIONAL: 'Project Plan 3',
  PROJECTPREMIUM: 'Project Plan 5',
  PROJECT_P1: 'Project Plan 1',
  VISIOCLIENT: 'Visio Plan 2',
  VISIO_PLAN1_DEPT: 'Visio Plan 1',
  EMS: 'Enterprise Mobility + Security E3',
  EMSPREMIUM: 'Enterprise Mobility + Security E5',
  AAD_PREMIUM: 'Microsoft Entra ID P1',
  AAD_PREMIUM_P2: 'Microsoft Entra ID P2',
  ATP_ENTERPRISE: 'Microsoft Defender for Office 365 (Plan 1)',
  THREAT_INTELLIGENCE: 'Microsoft Defender for Office 365 (Plan 2)',
  WIN_DEF_ATP: 'Microsoft Defender for Endpoint',
  IDENTITY_THREAT_PROTECTION: 'Microsoft 365 E5 Security',
  INTUNE_A: 'Microsoft Intune Plan 1',
  STREAM: 'Microsoft Stream',
  WINDOWS_STORE: 'Windows Store for Business',
  RIGHTSMANAGEMENT_ADHOC: 'Azure Rights Management (ad-hoc)',
  RIGHTSMANAGEMENT: 'Azure Information Protection Plan 1',
  MICROSOFT_BUSINESS_CENTER: 'Microsoft Business Center',
  WINDOWS_365_S_2_VCPU_4_GB_128_GB: 'Windows 365 Business',
}

// Markers that identify free/trial SKUs (kept separate from paid licenses).
const FREE_MARKERS = ['FREE', 'VIRAL', 'TRIAL', 'VTRIAL', 'EXPLORATORY', 'ADHOC', 'PRIVPREV']

export function skuFriendly(partNumber: string): string {
  if (SKU_NAMES[partNumber]) return SKU_NAMES[partNumber]
  // prettify: replace separators, title-case words, keep short acronyms upper
  return partNumber
    .replace(/[_\-.]+/g, ' ')
    .toLowerCase()
    .replace(/\b([a-z0-9]+)\b/g, (w) => (w.length <= 3 ? w.toUpperCase() : w[0].toUpperCase() + w.slice(1)))
    .trim()
}

// A SKU is treated as free/trial by name marker or by an effectively unlimited quota.
export function isFreeOrTrial(partNumber: string, total: number): boolean {
  const up = partNumber.toUpperCase()
  if (FREE_MARKERS.some((m) => up.includes(m))) return true
  return total >= 10000
}
