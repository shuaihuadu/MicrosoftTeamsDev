const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");
const config = require("./config");
const { Configuration, OpenAIApi } = require("openai");
const speechSdk = require("microsoft-cognitiveservices-speech-sdk");
const { BlobServiceClient } = require("@azure/storage-blob");
const { DefaultAzureCredential } = require("@azure/identity");
const uuidv1 = require('uuidv1');
const fs = require("fs");
const { cwd } = require("process");

const voiceNames = [
  { name: "普通话", value: "zh-CN-YunjianNeural" },
  {
    name: "粤语", value: "yue-CN-XiaoMinNeural"
  },
  {
    name: "河南话", value: "zh-CN-henan-YundengNeural"
  },
  {
    name: "东北话", value: "zh-CN-liaoning-XiaobeiNeural"
  }];

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    console.log(config);

    const configuration = new Configuration({
      apiKey: config.openaiApiKey
    });

    const openai = new OpenAIApi(configuration);

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      //console.log(context);
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Call the open ai to get response
      // const response = await openai.createCompletion({
      //   model: "text-davinci-003",
      //   prompt: txt,
      //   temperature: 0,
      //   max_tokens: 2048
      // });
      // console.log(response);
      // await context.sendActivity(response.data.choices[0].text);


      // 1.Process the open ai data result to speech
      // 2.Upload the file to Azure Blob Storage
      // 3.Send the file as a VoiceCard to teams
      const audioFile = "voice" + uuidv1() + ".wav";
      const speechConfig = speechSdk.SpeechConfig.fromSubscription(config.cognitiveServiceKey, config.cognitiveServiceRegion);
      const audioConfig = speechSdk.AudioConfig.fromAudioFileOutput(audioFile);
      speechConfig.speechSynthesisVoiceName = voiceNames[3].value;

      var synthesizer = new speechSdk.SpeechSynthesizer(speechConfig, audioConfig);
      synthesizer.speakTextAsync(txt,
        function (result) {
          if (result.reason === speechSdk.ResultReason.SynthesizingAudioCompleted) {
            //Upload the audio file to azure blob storage which convert by Azure Text To Speech
            const blobServiceClient = BlobServiceClient.fromConnectionString(config.azureStorageConnectionString);
            const containerClient = blobServiceClient.getContainerClient(config.azureStorageAccountContainerName);
            const blockBlobClient = containerClient.getBlockBlobClient(audioFile);
            var result = blockBlobClient.uploadFile(audioFile);
            console.log(blockBlobClient.url);
            //delete the audio file after 60s
            setTimeout(() => {
              fs.rm(audioFile);
            }, 60000);

            console.log("synthesis finished.");
          } else {
            console.error("Speech synthesis canceled, " + result.errorDetails +
              "\nDid you set the speech resource key and region values?");
          }
          synthesizer.close();
          synthesizer = null;
        },
        function (err) {
          console.trace("err - " + err);
          synthesizer.close();
          synthesizer = null;
        });
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          // const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          // await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          await context.sendActivity("我是一个基于OpenAI打造的智能机器人，你可以问我任何问题");
          break;
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;