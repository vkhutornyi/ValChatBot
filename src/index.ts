import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { CloudAdapter, ConfigurationBotFrameworkAuthentication, TurnContext } from "botbuilder";
import { ValChatBot } from "./bot";

const botAuth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: "MultiTenant",
});

const adapter = new CloudAdapter(botAuth);

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

    let statusCode = 200;
    let responseBody = "";

    const mockRes = {
      status(code: number) { statusCode = code; return this; },
      setHeader(_key: string, _value: string) { return this; },
      end(data?: string) { responseBody = data || ""; },
    };

    const mockReq = { body, headers, method: "POST" };

    try {
      await adapter.process(mockReq as any, mockRes as any, async (turnContext: TurnContext) => {
        await bot.run(turnContext);
      });
    } catch (err) {
      context.error("Adapter error:", err);
      return { status: 500, body: String(err) };
    }

    return { status: statusCode, body: responseBody };
  },
});
