import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessageFactory,
} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import OpenAI from "openai";
import config from "./config";

export class TeamsBot extends TeamsActivityHandler {
  private openai: OpenAI;

  constructor() {
    super();

    // Initialize OpenAI API
    this.openai = new OpenAI({ apiKey: config.openaiApiKey });

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
    
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }
    
      // Use OpenAI to generate a response
      const completion = await this.openai.chat.completions.create({
        messages: [{ role: "user", content: txt }],
        model: "gpt-4-1106-preview",
      });
    
      // Send the response back to the user as a reply to the specific message
      if (completion.choices && completion.choices.length > 0) {
        // await context.sendActivity(completion.choices[0].message.content.trim());
        const reply = MessageFactory.text(completion.choices[0].message.content.trim());
        reply.replyToId = context.activity.replyToId || context.activity.id;
        await context.sendActivity(reply);
      }
    
      await next();
    });
    

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }
}
