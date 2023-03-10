const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");
const config = require("./config");
const { Configuration, OpenAIApi } = require("openai");
const speechSdk = require("microsoft-cognitiveservices-speech-sdk");
const { BlobServiceClient } = require("@azure/storage-blob");
const uuidv1 = require('uuidv1');


const voiceNames = [
  { name: "普通话", value: "zh-CN-YunjianNeural" },
  { name: "粤语", value: "yue-CN-XiaoMinNeural" },
  { name: "河南话", value: "zh-CN-henan-YundengNeural" },
  { name: "东北话", value: "zh-CN-liaoning-XiaobeiNeural" }
];

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.audioFileObj = { url: "" };

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
      if (txt === "hello") {
        const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
        await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      } else {
        this.textToSpeech(txt, (audioFileUrl) => {
          this.audioFileObj.url = audioFileUrl;
          console.log(this.audioFileObj);
        });
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card1 = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          const card2 = cardTools.AdaptiveCards.declareWithoutData(rawLearnCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card1), CardFactory.adaptiveCard(card2)] });
          break;
        }
      }
      await next();
    });
  }


  textToSpeech(txt, callback) {
    // 1.Process the open ai data result to speech
    // 2.Upload the file to Azure Blob Storage
    // 3.Send the file url to a card in teams
    const audioFile = "voice" + uuidv1() + ".wav";
    const speechConfig = speechSdk.SpeechConfig.fromSubscription(config.cognitiveServiceKey, config.cognitiveServiceRegion);
    const audioConfig = speechSdk.AudioConfig.fromAudioFileOutput(audioFile);
    speechConfig.speechSynthesisVoiceName = voiceNames[0].value;

    var synthesizer = new speechSdk.SpeechSynthesizer(speechConfig, audioConfig);
    synthesizer.speakTextAsync(txt,
      function (result) {
        if (result.reason === speechSdk.ResultReason.SynthesizingAudioCompleted) {
          //Upload the audio file to azure blob storage which convert by Azure Text To Speech
          const blobServiceClient = BlobServiceClient.fromConnectionString(config.azureStorageConnectionString);
          const containerClient = blobServiceClient.getContainerClient(config.azureStorageAccountContainerName);
          const blockBlobClient = containerClient.getBlockBlobClient(audioFile);
          var result = blockBlobClient.uploadFile(audioFile);
          callback(blockBlobClient.url);
          console.log("synthesis finished.");
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
}

module.exports.TeamsBot = TeamsBot;