import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { BotFrameworkAdapter, TurnContext } from "botbuilder";
import { ValChatBot } from "./bot";

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId || "",
  appPassword: process.env.MicrosoftAppPassword || "",
});

adapter.onTurnError = async (context: TurnContext, error: Error) => {
  console.error("Bot turn error:", error);
  await context.sendActivity("The bot encountered an error. Please try again.");
};

const bot = new ValChatBot();

app.http("messages", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "messages",
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    context.log("Received message from Teams");

    const body = await req.text();
    const headers: Record<string, string> = {};
    req.headers.forEach((value, key) => { headers[key] = value; });

    return new Promise((resolve) => {
      const mockReq = { body, headers };

      const mockRes = {
        statusCode: 200,
        status: (code: number) => { mockRes.statusCode = code; return mockRes; },
        setHeader: (_key: string, _value: string) => { return mockRes; },
        end: (data?: string) => {
          resolve({ status: mockRes.statusCode, body: data || "" });
        },
        send: (data?: string) => {
          resolve({ status: mockRes.statusCode, body: data || "" });
        },
      };

      adapter.processActivity(mockReq as any, mockRes as any, async (turnContext: TurnContext) => {
        await bot.run(turnContext);
      }).catch((err) => {
        context.error("processActivity error:", err);
        resolve({ status: 500, body: String(err) });
      });
    });
  },
});
