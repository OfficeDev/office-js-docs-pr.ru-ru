---
title: Создание области задач Excel с помощью Vue
description: ''
ms.date: 09/04/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9947852a586570345ba9f3dfe09340af6d01ace6
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715634"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="0d029-102">Создание области задач Excel с помощью Vue</span><span class="sxs-lookup"><span data-stu-id="0d029-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="0d029-103">Из этой статьи вы узнаете, как создать надстройку области Excel, используя Vue и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="0d029-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0d029-104">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="0d029-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="0d029-105">Установите [Vue CLI](https://cli.vuejs.org/) глобально.</span><span class="sxs-lookup"><span data-stu-id="0d029-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="0d029-106">Создание нового приложения Vue</span><span class="sxs-lookup"><span data-stu-id="0d029-106">Generate a new Vue app</span></span>

<span data-ttu-id="0d029-p101">Используйте Vue CLI, чтобы создать новое приложение Vue. С помощью терминала выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="0d029-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command and then answer the prompts as described below.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="0d029-109">Затем выберите параметр `default`.</span><span class="sxs-lookup"><span data-stu-id="0d029-109">Then select the `default` preset.</span></span> <span data-ttu-id="0d029-110">Если в качестве пакета предлагается использовать Yarn или NPM, можно выбрать любой из этих вариантов.</span><span class="sxs-lookup"><span data-stu-id="0d029-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="0d029-111">Создание файла манифеста</span><span class="sxs-lookup"><span data-stu-id="0d029-111">Generate the manifest file</span></span>

<span data-ttu-id="0d029-112">У каждой надстройки должен быть файл манифеста, в нем определяются ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="0d029-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="0d029-113">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="0d029-113">Navigate to your app folder.</span></span>

   ```command&nbsp;line
   cd my-add-in
   ```

2. <span data-ttu-id="0d029-p103">Используя генератор Yeoman, создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="0d029-p103">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

   ```command&nbsp;line
   yo office
   ```

   ![Генератор Yeoman](../images/yo-office-manifest-only-vue.png)

   - <span data-ttu-id="0d029-117">**Выберите тип проекта:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="0d029-117">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
   - <span data-ttu-id="0d029-118">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="0d029-118">**What do you want to name your add-in?**</span></span> `my-office-add-in`
   - <span data-ttu-id="0d029-119">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="0d029-119">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="0d029-120">После завершения работы мастера создается папка `my-office-add-in`, содержащая файл `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="0d029-120">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="0d029-121">В конце краткого руководства вам потребуется использовать манифест для загрузки без публикации и тестирования вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="0d029-121">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="0d029-122">Защита приложения</span><span class="sxs-lookup"><span data-stu-id="0d029-122">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="0d029-123">Чтобы включить HTTPS для своего приложения, создайте файл `vue.config.js` в корневой папке проекта Vue со следующим содержанием:</span><span class="sxs-lookup"><span data-stu-id="0d029-123">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="0d029-124">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="0d029-124">Update the app</span></span>

1. <span data-ttu-id="0d029-125">Откройте файл `public/index.html` и добавьте следующий тег `<script>` непосредственно перед тегом `</head>`:</span><span class="sxs-lookup"><span data-stu-id="0d029-125">Open `public/index.html`, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="0d029-126">Откройте файл `src/main.js` и замените его содержимое указанным ниже кодом:</span><span class="sxs-lookup"><span data-stu-id="0d029-126">Open the `src/main.js` file and replace its contents with the following code:</span></span>

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. <span data-ttu-id="0d029-127">Откройте файл `src/App.vue` и замените содержимое файла указанным ниже кодом:</span><span class="sxs-lookup"><span data-stu-id="0d029-127">Open the `src/App.vue` file and replace its contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
       background: #fff;
       position: fixed;
       top: 80px;
       left: 0;
       right: 0;
       bottom: 0;
       overflow: auto;
     }

     .padding {
       padding: 15px;
     }
   </style>
   ```

## <a name="start-the-dev-server"></a><span data-ttu-id="0d029-128">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="0d029-128">Start the dev server</span></span>

1. <span data-ttu-id="0d029-129">Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="0d029-129">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="0d029-130">В веб-браузере перейдите по адресу `https://localhost:3000` (обратите внимание на `https`).</span><span class="sxs-lookup"><span data-stu-id="0d029-130">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="0d029-131">Если появится сообщение, что сертификат сайта не является доверенным, [сделайте так, чтобы компьютер ему доверял](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="0d029-131">If your browser indicates that the site's certificate is not trusted, you will need to configure your computer to trust the certificate.</span></span>

3. <span data-ttu-id="0d029-132">Если страница на сайте `https://localhost:3000` пуста, при этом нет ошибок сертификата, значит она работает.</span><span class="sxs-lookup"><span data-stu-id="0d029-132">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="0d029-133">Приложение Vue подключается после запуска Office, поэтому в нем отображаются только элементы из среды Excel.</span><span class="sxs-lookup"><span data-stu-id="0d029-133">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="0d029-134">Проверка</span><span class="sxs-lookup"><span data-stu-id="0d029-134">Try it out</span></span>

1. <span data-ttu-id="0d029-135">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="0d029-135">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="0d029-136">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="0d029-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="0d029-137">Веб-браузер: [загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="0d029-137">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="0d029-138">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="0d029-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="0d029-139">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="0d029-139">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="0d029-141">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="0d029-141">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="0d029-142">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="0d029-142">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="0d029-144">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0d029-144">Next steps</span></span>

<span data-ttu-id="0d029-145">Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Vue!</span><span class="sxs-lookup"><span data-stu-id="0d029-145">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="0d029-146">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="0d029-146">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0d029-147">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="0d029-147">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="0d029-148">См. также</span><span class="sxs-lookup"><span data-stu-id="0d029-148">See also</span></span>

* [<span data-ttu-id="0d029-149">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="0d029-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="0d029-150">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0d029-150">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="0d029-151">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="0d029-151">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="0d029-152">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0d029-152">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
