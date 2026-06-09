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
  route: "api/messages",
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    context.log("Received message from Teams");

    const body = await req.text();
    const headers: Record<string, string> = {};
    req.headers.forEach((value, key) => { headers[key] = value; });

    await adapter.process(
      { body, headers, method: "POST" } as any,
      {} as any,
      async (turnContext: TurnContext) => {
        await bot.run(turnContext);
      }
    );

    return { status: 200 };
  },
});
