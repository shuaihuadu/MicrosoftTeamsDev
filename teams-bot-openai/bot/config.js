const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  openaiApiKey: process.env.OPENAI_API_KEY,
  cognitiveServiceKey: process.env.COGNITIVE_SERVICE_KEY,
  cognitiveServiceRegion: process.env.COGNITIVE_SERVICE_REGION,
  cognitiveServiceEndpoint: process.env.COGNITIVE_SERVICE_ENDPOINT,
  azureStorageAccountName: process.env.AZURE_STORAGE_ACCOUNT_NAME,
  azureStorageAccountContainerName: process.env.AZURE_STORAGE_ACCOUNT_CONTAINER_NAME
};

module.exports = config;