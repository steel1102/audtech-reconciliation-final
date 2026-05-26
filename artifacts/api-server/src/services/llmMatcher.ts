import { GoogleGenerativeAI } from "@google/generative-ai";

const genAI = new GoogleGenerativeAI(
  process.env.GEMINI_API_KEY || ""
);

const model = genAI.getGenerativeModel({
  model: "gemini-1.5-flash-latest",
});

export async function classifyLedgerName(
  ledgerName: string
): Promise<string | null> {
  try {
    console.log("LLM CALLED:", ledgerName);

    const prompt = `
You are a financial reconciliation assistant.

Classify this ledger/account name into a short financial category.

Ledger:
"${ledgerName}"

Examples:
- Trade Debtors → receivable
- Salary Payable → payable
- Raw Material Inventory → inventory
- Bank Charges → expense
- GST Receivable → tax receivable

Return ONLY the category name.
`;

    const result = await model.generateContent(prompt);

    const response = result.response.text();

    console.log("LLM RESULT:", ledgerName, response);

    return response;
  } catch (error) {
    console.error("LLM classification failed:", error);

    return null;
  }
}