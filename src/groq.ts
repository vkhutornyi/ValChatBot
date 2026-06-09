import Groq from "groq-sdk";

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });
const model = process.env.GROQ_MODEL || "llama-3.3-70b-versatile";

const SYSTEM_PROMPT = `You are ValChatBot, an AI assistant integrated into Microsoft Teams.
You help team members with questions, tasks, and information.
Be concise, helpful, and professional. You can respond in the same language the user writes in.
When responding in Teams, keep messages clear and well-structured.`;

export async function askGroq(userMessage: string, conversationHistory: Array<{ role: "user" | "assistant"; content: string }> = []): Promise<string> {
  const messages: Array<{ role: "system" | "user" | "assistant"; content: string }> = [
    { role: "system", content: SYSTEM_PROMPT },
    ...conversationHistory,
    { role: "user", content: userMessage },
  ];

  const response = await groq.chat.completions.create({
    model,
    messages,
    max_tokens: 1024,
    temperature: 0.7,
  });

  return response.choices[0]?.message?.content || "Sorry, I could not generate a response.";
}
