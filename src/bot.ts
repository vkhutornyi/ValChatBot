import { ActivityHandler, MessageFactory, TurnContext } from "botbuilder";
import { askGroq } from "./groq";

export class ValChatBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context: TurnContext, next) => {
      const activity = context.activity;
      const rawText = activity.text || "";

      // Strip HTML tags and @mentions that Teams injects
      const userMessage = rawText
        .replace(/<[^>]*>/g, "")
        .replace(/@ValChatBot/gi, "")
        .trim();

      if (!userMessage) {
        await next();
        return;
      }

      // Show typing indicator while generating response
      await context.sendActivity({ type: "typing" });

      try {
        const reply = await askGroq(userMessage);
        await context.sendActivity(MessageFactory.text(reply));
      } catch (error) {
        console.error("Groq API error:", error);
        await context.sendActivity(MessageFactory.text("⚠️ Sorry, I encountered an error. Please try again."));
      }

      await next();
    });

    this.onMembersAdded(async (context: TurnContext, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(MessageFactory.text("👋 Hi! I'm **ValChatBot**, your AI assistant. Mention me with @ValChatBot and ask me anything!"));
        }
      }
      await next();
    });
  }
}
