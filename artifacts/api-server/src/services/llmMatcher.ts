const ACCOUNT_ALIASES: Record<string, string[]> = {
  "Inventory": [
    "inventory",
    "raw material",
    "stock",
    "closing stock"
  ],

  "Trade Receivables": [
    "debtor",
    "trade debtor",
    "receivable",
    "sundry debtor"
  ],

  "Trade Payables": [
    "creditor",
    "trade creditor",
    "payable",
    "sundry creditor"
  ],

  "Cash and Bank": [
    "cash",
    "bank",
    "hdfc",
    "icici",
    "sbi"
  ],

  "Salary Payable": [
    "salary payable",
    "outstanding salary",
    "salary outstanding"
  ],

  "Fixed Assets": [
    "plant",
    "machinery",
    "furniture",
    "vehicle",
    "equipment"
  ]
};

function normalize(text: string): string {
  return text.toLowerCase().trim();
}

export async function classifyLedgerName(
  ledgerName: string
): Promise<string | null> {

  const normalized = normalize(ledgerName);

  for (const [category, aliases] of Object.entries(ACCOUNT_ALIASES)) {

    for (const alias of aliases) {

      if (normalized.includes(normalize(alias))) {

        console.log("LOCAL LLM MATCH:", ledgerName, "→", category);

        return category;
      }
    }
  }

  console.log("NO LOCAL MATCH:", ledgerName);

  return ledgerName;
}