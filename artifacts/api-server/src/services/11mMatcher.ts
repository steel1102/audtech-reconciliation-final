import OpenAI from "openai";

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function classifyLedgerName(name: string) {
  try {
    const prompt = `
You are a financial reconciliation assistant.

Determine whether this line item is:
1. A real ledger/account item
2. A report/narrative/non-ledger item

Return ONLY JSON.

Examples:
"Trade Receivables" -> ledger
"Cash and Cash Equivalents" -> ledger
"Director Report" -> non-ledger
"Statement of Financial Position" -> non-ledger

Input:
${name}
`;
console.log("LLM CALLED:", name);
    const response = await client.chat.completions.create({
      model: "gpt-4.1-mini",
      messages: [
        {
          role: "user",
          content: prompt,
        },
      ],
      temperature: 0,
    });

    return response.choices[0].message.content;
  } catch (err) {
    console.error("LLM classification failed:", err);
    return null;
  }
}