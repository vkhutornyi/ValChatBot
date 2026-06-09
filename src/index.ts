import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { BotFrameworkAdapter } from "botbuilder";
import { ValChatBot } from "./bot";

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});

adapter.onTurnError = async (context, error) => {
  console.error("Bot turn error:", error);
  await context.sendTraceActivity("OnTurnError", `${error}`, "https://www.botframework.com/schemas/error", "TurnError");
  await context.sendActivity("The bot encountered an error. Please try again.");
};

const bot = new ValChatBot();

app.http("messages", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "api/messages",
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const body = await req.text();

    return new Promise((resolve) => {
      const mockReq = {
        body: Buffer.from(body),
        headers: Object.fromEntries(req.headers.entries()),
      };

      const mockRes = {
        statusCode: 200,
        body: "",
        setHeader: () => {},
        end: (data?: string) => {
          resolve({ status: mockRes.statusCode, body: data || mockRes.body });
        },
      };

      adapter.processActivity(mockReq as any, mockRes as any, async (turnContext) => {
        await bot.run(turnContext);
      }).catch((err) => {
        context.error("processActivity error:", err);
        resolve({ status: 500, body: "Internal Server Error" });
      });
    });
  },
});
