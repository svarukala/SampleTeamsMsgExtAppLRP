const { TeamsActivityHandler, CardFactory, MessageFactory, TeamsInfo } = require("botbuilder");
const config = require("./config");

class ActionApp extends TeamsActivityHandler {

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const attachment = CardFactory.adaptiveCard({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `${data.title}`,
          wrap: true,
          size: "Large",
        },
        {
          type: "TextBlock",
          text: `${data.subTitle}`,
          wrap: true,
          size: "Medium",
        },
        {
          type: "TextBlock",
          text: `${data.text}`,
          wrap: true,
          size: "Small",
        },
      ],
    });

    // Send the card to a Teams channel as a proactive message
    try {
      const teamsChannelId = "19:MLVoXzCRR_GtpqSWyOFePxl8Wj8DdKVq9PHiY4PhTZc1@thread.tacv2";
      const activity = MessageFactory.attachment(attachment);
      //MessageFactory.text('This will be the first message in a new thread');
      const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, config.botId);      
    } catch (error) {
      console.log(error);      
    }

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
  }

  /*
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    const data = action.data;
    const attachment = CardFactory.adaptiveCard({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: `${data.title}`,
          wrap: true,
          size: "Large",
        },
        {
          type: "TextBlock",
          text: `${data.subTitle}`,
          wrap: true,
          size: "Medium",
        },
        {
          type: "TextBlock",
          text: `${data.text}`,
          wrap: true,
          size: "Small",
        },
      ],
    });

    const teamsChannelId = "19:MLVoXzCRR_GtpqSWyOFePxl8Wj8DdKVq9PHiY4PhTZc1@thread.tacv2";
    const activity = MessageFactory.text('This will be the first message in a new thread');
    const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, config.botId);


    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
  }
  */
}
module.exports.ActionApp = ActionApp;
