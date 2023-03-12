const { TeamsActivityHandler, CardFactory, TurnContext, MessageFactory } = require("botbuilder");
const { ActionTypes } = require('botframework-schema');
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawAudioCard = require("./adaptiveCards/audio.json");
const cardTools = require("@microsoft/adaptivecards-tools");
const config = require("./config");
const { Configuration, OpenAIApi } = require("openai");
const speechSdk = require("microsoft-cognitiveservices-speech-sdk");
const { BlobServiceClient } = require("@azure/storage-blob");
const uuidv1 = require('uuidv1');
const lodash = require("lodash");

const voiceNames = [
  { type: ActionTypes.ImBack, title: "普通话", value: "zh-CN-YunjianNeural" },
  { type: ActionTypes.ImBack, title: "粤语", value: "yue-CN-XiaoMinNeural" },
  { type: ActionTypes.ImBack, title: "河南话", value: "zh-CN-henan-YundengNeural" }
  //{ type: ActionTypes.ImBack, title: "东北话", value: "zh-CN-liaoning-XiaobeiNeural" },
  //Suggested Actions only show there actions on Teams
];

const voiceNameTitles = lodash.map(voiceNames, "title");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.chatData = { audioFlileUrl: "", voiceName: voiceNames[0], lastGPTAnswer: "" };
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }
      if (txt === "hello") {
        const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      } else {
        if (voiceNameTitles.includes(txt)) {
          var voiceName = lodash.find(voiceNames, function (v) {
            return v.title === txt;
          });
          if (!voiceName) {
            voiceName = voiceNames[0];
          }
          const audioFile = ["voice", uuidv1(), ".wav"].join("");
          this.textToSpeech(audioFile);
          this.chatData.voiceName = voiceName;
          this.chatData.audioFlileUrl = [config.azureStorageAccountContainerUrl, audioFile].join("/");
          const card = cardTools.AdaptiveCards.declare(rawAudioCard).render(this.chatData);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
        }
        else {
          try {
            const configuration = new Configuration({
              apiKey: config.openaiApiKey
            });
            const openai = new OpenAIApi(configuration);
            // Send request by the open ai sdk and get response
            const completion = await openai.createCompletion({
              model: "text-davinci-003",
              prompt: txt,
              temperature: 0,
              max_tokens: 2048//The response text max length
            });

            this.chatData.lastGPTAnswer = completion.data.choices[0].text;
            console.log(this.chatData);
            await context.sendActivity(this.chatData.lastGPTAnswer);
            await this.sendSuggestedActions(context);
          } catch (error) {
            await context.sendActivity("error:" + JSON.stringify(error));
            await this.sendSuggestedActions(context);
          }
        }
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }
  textToSpeech(audioFile) {
    // 1.Process the open ai data result to speech
    // 2.Upload the file to Azure Blob Storage
    // 3.Send the file url to a card in teams
    const speechConfig = speechSdk.SpeechConfig.fromSubscription(config.cognitiveServiceKey, config.cognitiveServiceRegion);
    const audioConfig = speechSdk.AudioConfig.fromAudioFileOutput(audioFile);
    if (!this.chatData.voiceName) {
      this.chatData.voiceName = voiceNames;
    }
    speechConfig.speechSynthesisVoiceName = this.chatData.voiceName.value;

    var synthesizer = new speechSdk.SpeechSynthesizer(speechConfig, audioConfig);
    synthesizer.speakTextAsync(this.chatData.lastGPTAnswer, (result) => {
      if (result.reason === speechSdk.ResultReason.SynthesizingAudioCompleted) {
        //Upload the audio file to azure blob storage which convert by Azure Text To Speech
        const blobServiceClient = BlobServiceClient.fromConnectionString(config.azureStorageConnectionString);
        const containerClient = blobServiceClient.getContainerClient(config.azureStorageAccountContainerName);
        const blockBlobClient = containerClient.getBlockBlobClient(audioFile);
        var result = blockBlobClient.uploadFile(audioFile);
        this.chatData.audioFlileUrl = blockBlobClient.url;
        //delete the audio file after 60s
        // setTimeout(() => {
        //   fs.rm(audioFile);
        // }, 60000);
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
  }
  async sendSuggestedActions(context) {
    var reply = MessageFactory.text("我还可以将我的回答转换成语音，您如果需要的话可以点击下面的语音选项，我会按照您的选择进行转换。");
    reply.suggestedActions = { "actions": voiceNames, "to": [context.activity.from.id] };
    await context.sendActivity(reply);
  }
}

module.exports.TeamsBot = TeamsBot;