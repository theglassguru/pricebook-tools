/**
 * FieldPulse Pricebook 2.0 — Calculated Column Registry
 *
 * Canonical source of truth for every ARRAYFORMULA on the Items,
 * Item Option Values, and Item Groupings sheets.
 *
 * Consumed by repairCalculatedColumns() in Menu.gs.
 *
 * Each entry maps a column letter to its anchor-row-2 ARRAYFORMULA.
 * Repair clears the column body (row 2 → maxRow) and writes the formula
 * back at row 2; the ARRAYFORMULA spills the rest.
 *
 * Flat-fee pass-through branch (Service surcharges, etc.):
 *   Items: flat = (Type="SERVICE") * (Markup=1)
 *   Item Option Values: flat = (Markup=1)
 *   When flat, Energy Surcharge = 0, Landed Cost = 0,
 *   Price = Cost, Net Markup = 1.
 */

const CALC_SHEETS = [
  // ===========================================================================
  // ITEMS — 26 ARRAYFORMULAs
  // ===========================================================================
  {
    name: 'Items',
    displayLabel: 'Items',
    headerRow: 1,
    formulas: {
      J: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, typ,F2:F,
  rate,Variables!$C$4,
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (UPPER(TRIM(typ))<>"SERVICE"),
  IF(cost="","",cost*rate*gate)
))`,
      K: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, markup,I2:I, typ,F2:F,
  rate,Variables!$C$4, mode,Variables!$C$5,
  flat,(UPPER(TRIM(typ))="SERVICE")*(markup=1),
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (UPPER(TRIM(typ))<>"SERVICE"),
  surcharge,cost*rate*gate,
  IF(cost="","",IF(flat,0,cost+IF(mode="Markup",surcharge,0)))
))`,
      L: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, markup,I2:I, typ,F2:F,
  rate,Variables!$C$4, mode,Variables!$C$5,
  flat,(UPPER(TRIM(typ))="SERVICE")*(markup=1),
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (UPPER(TRIM(typ))<>"SERVICE"),
  surcharge,cost*rate*gate,
  landed,cost+IF(mode="Markup",surcharge,0),
  IF((cost="")+(markup="")>0,"",IF(flat,cost,landed*markup+IF(mode="Markup",0,surcharge)))
))`,
      M: `=ARRAYFORMULA(LET(
  cost,H2:H, price,L2:L, markup,I2:I, typ,F2:F,
  flat,(UPPER(TRIM(typ))="SERVICE")*(markup=1),
  IF(flat,
    IF(cost="","",1),
    IF((cost="")+(price="")+(IFERROR(cost*1,0)=0)>0,"",price/cost))
))`,
      Z:  `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),REGEXREPLACE(TRIM(E2:E),"\\s+"," "),""))`,
      AA: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(ISNUMBER(SEARCH("HARDWARE",D2:D)),REGEXEXTRACT(D2:D&"","^[^,]+")&"-HDWR",SUBSTITUTE(D2:D,", ","-")&IF((B2:B<>"")*(C2:C<>""),"-"&SUBSTITUTE(B2:B,CHAR(34),"")&"-"&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(C2:C,"Patterned / Textured, Annealed","PT-A"),"Patterned / Textured, Tempered","PT-T"),"Reflective / Tinted, Annealed","RT-A"),"Reflective / Tinted, Tempered","RT-T"),"Laminated, Annealed","IMPACT"),"Annealed","A"),"Tempered","T"),"Laminated","L"),"Mirror","M"),IF(TRIM(REGEXREPLACE(E2:E,"^(IGU|Monolithic Glass|Flat Mirror|Glass Hardware|Mirror Hardware|Glass|Mirror|Doors?|Screens?|Skylights?|Windows?)\\s*",""))="","","-"&REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(UPPER(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(REGEXREPLACE(E2:E,"^(IGU|Monolithic Glass|Flat Mirror|Glass Hardware|Mirror Hardware|Glass|Mirror|Doors?|Screens?|Skylights?|Windows?)\\s*",""))," / ","-"),"(",""),")",""),CHAR(34),"")," ","-"),"-+","-")),"-SINGLE$","-S"),"-DOUBLE$","-D"),"-TRIPLE$","-T")))),""))`,
      AB: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),LOWER(TRIM(F2:F)),""))`,
      AC: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),0,""))`,
      AD: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),0,""))`,
      AE: `=ARRAYFORMULA(IF((A2:A=TRUE)*(N2:N<>""),SUBSTITUTE(N2:N,", ",","),""))`,
      AF: `=ARRAYFORMULA(IF((A2:A=TRUE)*(G2:G<>""),TRIM(G2:G),""))`,
      AG: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(K2:K="",0,K2:K),""))`,
      AH: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(L2:L="",0,L2:L),""))`,
      AI: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(ISNUMBER(SEARCH("IGU",D2:D)),VLOOKUP("Min. IGU Size Charged",Variables!A:C,3,FALSE),IF((D2:D="GLASS")+(D2:D="MIRROR")>0,VLOOKUP("Min. Glass Size Charged",Variables!A:C,3,FALSE),"")),""))`,
      AJ: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(ISNUMBER(SEARCH("IGU",D2:D)),IFERROR(VLOOKUP("Max IGU Size "&B2:B&" "&IF(ISNUMBER(SEARCH("Tempered",C2:C)),"T","A"),Variables!A:C,3,FALSE),""),IF((D2:D="GLASS")+(D2:D="MIRROR")>0,IFERROR(VLOOKUP("Max GLASS Size "&B2:B&" "&IF(ISNUMBER(SEARCH("Tempered",C2:C)),"T","A"),Variables!A:C,3,FALSE),""),"")),""))`,
      AN: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(TRIM(F2:F)="Product","Yes",IF(TRIM(F2:F)="Service","No","")),""))`,
      AO: `=ARRAYFORMULA(IF((A2:A=TRUE)*(O2:O<>""),O2:O,""))`,
      AP: `=ARRAYFORMULA(IF((A2:A=TRUE)*(P2:P<>""),P2:P,""))`,
      AQ: `=ARRAYFORMULA(IF((A2:A=TRUE)*(Q2:Q<>""),Q2:Q,""))`,
      AR: `=ARRAYFORMULA(IF((A2:A=TRUE)*(R2:R<>""),R2:R,""))`,
      AS: `=ARRAYFORMULA(IF((A2:A=TRUE)*(S2:S<>""),S2:S,""))`,
      AT: `=ARRAYFORMULA(IF((A2:A=TRUE)*(T2:T<>""),T2:T,""))`,
      AU: `=ARRAYFORMULA(IF((A2:A=TRUE)*(U2:U<>""),U2:U,""))`,
      AV: `=ARRAYFORMULA(IF((A2:A=TRUE)*(V2:V<>""),V2:V,""))`,
      AW: `=ARRAYFORMULA(IF((A2:A=TRUE)*(W2:W<>""),W2:W,""))`
    }
  },

  // ===========================================================================
  // ITEM OPTION VALUES — 9 ARRAYFORMULAs
  // ===========================================================================
  {
    name: 'Item Option Values',
    displayLabel: 'Item Option Values',
    headerRow: 1,
    formulas: {
      G: `=ARRAYFORMULA(IF(E2:E="","",IFERROR(IF(VLOOKUP(E2:E,item_option_names,3,FALSE)="Flat Pricing","Flat Pricing",VLOOKUP(E2:E,item_option_names,4,FALSE)),"")))`,
      J: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, markup,I2:I,
  rate,Variables!$C$4,
  flat,(markup=1),
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (1-flat),
  IF(cost="","",cost*rate*gate)
))`,
      K: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, markup,I2:I,
  rate,Variables!$C$4, mode,Variables!$C$5,
  flat,(markup=1),
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (1-flat),
  surcharge,cost*rate*gate,
  IF(cost="","",IF(flat,0,cost+IF(mode="Markup",surcharge,0)))
))`,
      L: `=ARRAYFORMULA(LET(
  cat,D2:D, cost,H2:H, markup,I2:I,
  rate,Variables!$C$4, mode,Variables!$C$5,
  flat,(markup=1),
  gate,(cat<>"")*
       ((ISNUMBER(SEARCH("GLASS",cat))+ISNUMBER(SEARCH("IGU",cat))+ISNUMBER(SEARCH("MIRROR",cat)))>0)*
       ((ISNUMBER(SEARCH("HARDWARE",cat))+ISNUMBER(SEARCH("OTHER",cat)))=0)*
       (1-flat),
  surcharge,cost*rate*gate,
  landed,cost+IF(mode="Markup",surcharge,0),
  IF((cost="")+(markup="")>0,"",IF(flat,cost,landed*markup+IF(mode="Markup",0,surcharge)))
))`,
      M: `=ARRAYFORMULA(LET(
  cost,H2:H, price,L2:L, markup,I2:I,
  flat,(markup=1),
  IF(flat,
    IF(cost="","",1),
    IF((cost="")+(price="")+(IFERROR(cost*1,0)=0)>0,"",price/cost))
))`,
      N: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),E2:E,""))`,
      O: `=ARRAYFORMULA(IF((A2:A=TRUE)*(F2:F<>""),F2:F,""))`,
      P: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(K2:K="",0,K2:K),0))`,
      Q: `=ARRAYFORMULA(IF((A2:A=TRUE)*(E2:E<>""),IF(L2:L="",0,L2:L),""))`
    }
  },

  // ===========================================================================
  // ITEM GROUPINGS — 31 ARRAYFORMULAs
  // ===========================================================================
  {
    name: 'Item Groupings',
    displayLabel: 'Item Groupings',
    headerRow: 1,
    formulas: {
      D: `=ARRAYFORMULA(IF((C2:C<>"")*(C2:C<>0), "Manual Markup", ""))`,
      E: `=ARRAYFORMULA(IF((C2:C<>"")*(C2:C<>0), 0, ""))`,
      G: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 16, FALSE), "")))`,
      H: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 2, FALSE), "")))`,
      I: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 4, FALSE), "")))`,
      J: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 5, FALSE), "")))`,
      K: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 15, FALSE), "")))`,
      L: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 13, FALSE), "")))`,
      M: `=ARRAYFORMULA(IF($F2:$F="", "", IFERROR(VLOOKUP($F2:$F, items, 3, FALSE), "")))`,
      X:  `=ARRAYFORMULA(IF($A2:$A=TRUE, B2:B, IFERROR(1/0)))`,
      Y:  `=ARRAYFORMULA(IF($A2:$A=TRUE, C2:C, IFERROR(1/0)))`,
      Z:  `=ARRAYFORMULA(IF($A2:$A=TRUE, D2:D, IFERROR(1/0)))`,
      AA: `=ARRAYFORMULA(IF($A2:$A=TRUE, E2:E, IFERROR(1/0)))`,
      AB: `=ARRAYFORMULA(IF($A2:$A=TRUE, F2:F, IFERROR(1/0)))`,
      AC: `=ARRAYFORMULA(IF($A2:$A=TRUE, G2:G, IFERROR(1/0)))`,
      AD: `=ARRAYFORMULA(IF($A2:$A=TRUE, H2:H, IFERROR(1/0)))`,
      AE: `=ARRAYFORMULA(IF($A2:$A=TRUE, I2:I, IFERROR(1/0)))`,
      AF: `=ARRAYFORMULA(IF($A2:$A=TRUE, J2:J, IFERROR(1/0)))`,
      AG: `=ARRAYFORMULA(IF($A2:$A=TRUE, K2:K, IFERROR(1/0)))`,
      AH: `=ARRAYFORMULA(IF($A2:$A=TRUE, L2:L, IFERROR(1/0)))`,
      AI: `=ARRAYFORMULA(IF($A2:$A=TRUE, M2:M, IFERROR(1/0)))`,
      AJ: `=ARRAYFORMULA(IF($A2:$A=TRUE, N2:N, IFERROR(1/0)))`,
      AK: `=ARRAYFORMULA(IF($A2:$A=TRUE, O2:O, IFERROR(1/0)))`,
      AL: `=ARRAYFORMULA(IF($A2:$A=TRUE, P2:P, IFERROR(1/0)))`,
      AM: `=ARRAYFORMULA(IF($A2:$A=TRUE, Q2:Q, IFERROR(1/0)))`,
      AN: `=ARRAYFORMULA(IF($A2:$A=TRUE, R2:R, IFERROR(1/0)))`,
      AO: `=ARRAYFORMULA(IF($A2:$A=TRUE, S2:S, IFERROR(1/0)))`,
      AP: `=ARRAYFORMULA(IF($A2:$A=TRUE, T2:T, IFERROR(1/0)))`,
      AQ: `=ARRAYFORMULA(IF($A2:$A=TRUE, U2:U, IFERROR(1/0)))`,
      AR: `=ARRAYFORMULA(IF($A2:$A=TRUE, V2:V, IFERROR(1/0)))`,
      AS: `=ARRAYFORMULA(IF($A2:$A=TRUE, W2:W, IFERROR(1/0)))`
    }
  }
];

function colLetterToIndex_(letter) {
  let result = 0;
  const upper = String(letter).toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result;
}

function calcSheetsTotalFormulas_() {
  let total = 0;
  for (let i = 0; i < CALC_SHEETS.length; i++) {
    total += Object.keys(CALC_SHEETS[i].formulas).length;
  }
  return total;
}
// clasp+git pipeline verified 2026-05-05
