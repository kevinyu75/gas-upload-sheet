function runGemini() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("gemini_api_key");
  const model = "gemini-1.5-pro"; // 這裡直接指定模型
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const prompt = `
你是一個專業的電商文案助手，請幫我生成一段大約 500 字的商品介紹文案。
風格要求：親切、易讀、口語化、像跟朋友聊天一樣，不能有強烈推銷感，但要能吸引購買。
商品名稱：紅龍牛肉湯
商品描述：在家也能享用百元牛肉麵
紅龍牛肉湯 $65/包
吃過必定回購 紅龍紅燒牛肉湯
在家自己加上一團麵一樣也可以吃到道地牛肉麵🍜
一包紅龍牛肉湯加點麵加點青菜
就是熱銷的美食好好吃😋
看的見肉塊的紅龍紅燒牛肉湯！
香濃入味湯頭溫醇
軟嫩的肉質彈性有嚼勁
宅在家一個人也好方便
正餐宵夜全家人也都搶著吃🤤
喜歡就多囤幾包想吃就可以快速料理哦
團購價: $65/包
  `;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  Logger.log(result.candidates[0].content.parts[0].text);
}
