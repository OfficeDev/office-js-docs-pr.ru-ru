---
title: Создание области задач Excel с помощью Vue
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и Vue.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 41448463f33edf7bdddd27981db4e427dd0fcb74
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950735"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="17fd9-103">Создание области задач Excel с помощью Vue</span><span class="sxs-lookup"><span data-stu-id="17fd9-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="17fd9-104">Из этой статьи вы узнаете, как создать надстройку области Excel, используя Vue и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="17fd9-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="17fd9-105">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="17fd9-105">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="17fd9-106">Установите [Vue CLI](https://cli.vuejs.org/) глобально.</span><span class="sxs-lookup"><span data-stu-id="17fd9-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="17fd9-107">Создание нового приложения Vue</span><span class="sxs-lookup"><span data-stu-id="17fd9-107">Generate a new Vue app</span></span>

<span data-ttu-id="17fd9-p101">Используйте Vue CLI, чтобы создать новое приложение Vue. С помощью терминала выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="17fd9-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="17fd9-110">Затем выберите параметр `default`.</span><span class="sxs-lookup"><span data-stu-id="17fd9-110">Then select the `default` preset.</span></span> <span data-ttu-id="17fd9-111">Если в качестве пакета предлагается использовать Yarn или NPM, можно выбрать любой из этих вариантов.</span><span class="sxs-lookup"><span data-stu-id="17fd9-111">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="17fd9-112">Создание файла манифеста</span><span class="sxs-lookup"><span data-stu-id="17fd9-112">Generate the manifest file</span></span>

<span data-ttu-id="17fd9-113">У каждой надстройки должен быть файл манифеста, в нем определяются ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="17fd9-113">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="17fd9-114">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="17fd9-114">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="17fd9-115">С помощью генератора Yeoman создайте файл манифеста для надстройки, выполнив следующую команду:</span><span class="sxs-lookup"><span data-stu-id="17fd9-115">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="17fd9-116">При выполнении команды `yo office` может появиться запрос о политиках сбора данных генератора Yeoman и средств CLI надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="17fd9-116">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="17fd9-117">Используйте предоставленные сведения, чтобы ответить на запросы подходящим образом.</span><span class="sxs-lookup"><span data-stu-id="17fd9-117">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="17fd9-118">Если в ответ на второй запрос выбран вариант **Выход**, потребуется снова выполнить команду `yo office`, когда вы будете готовы создать проект надстройки.</span><span class="sxs-lookup"><span data-stu-id="17fd9-118">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="17fd9-119">При появлении запроса предоставьте следующую информацию для создания проекта надстройки:</span><span class="sxs-lookup"><span data-stu-id="17fd9-119">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="17fd9-120">**Выберите тип проекта:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="17fd9-120">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="17fd9-121">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="17fd9-121">**What do you want to name your add-in?**</span></span> `my-office-add-in`
    - <span data-ttu-id="17fd9-122">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="17fd9-122">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Генератор Yeoman](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="17fd9-124">После завершения работы мастера создается папка `my-office-add-in`, содержащая файл `manifest.xml`.</span><span class="sxs-lookup"><span data-stu-id="17fd9-124">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="17fd9-125">В конце краткого руководства вам потребуется использовать манифест для загрузки без публикации и тестирования вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="17fd9-125">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="17fd9-126">Вы можете игнорировать инструкции по *дальнейшим действиям*, предоставляемые генератором Yeoman после создания проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="17fd9-126">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="17fd9-127">Пошаговые инструкции этой статьи содержат все сведения, необходимые для завершения этого учебного курса.</span><span class="sxs-lookup"><span data-stu-id="17fd9-127">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="17fd9-128">Защита приложения</span><span class="sxs-lookup"><span data-stu-id="17fd9-128">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="17fd9-129">Чтобы включить HTTPS для своего приложения, создайте файл `vue.config.js` в корневой папке проекта Vue со следующим содержанием:</span><span class="sxs-lookup"><span data-stu-id="17fd9-129">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="17fd9-130">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="17fd9-130">Update the app</span></span>

1. <span data-ttu-id="17fd9-131">Откройте файл `public/index.html` и добавьте следующий тег `<script>` непосредственно перед тегом `</head>`:</span><span class="sxs-lookup"><span data-stu-id="17fd9-131">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="17fd9-132">Откройте файл `src/main.js` и замените его содержимое указанным ниже кодом:</span><span class="sxs-lookup"><span data-stu-id="17fd9-132">Open `src/main.js` and replace the contents with the following code:</span></span>

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

3. <span data-ttu-id="17fd9-133">Откройте файл `src/App.vue` и замените содержимое файла указанным ниже кодом:</span><span class="sxs-lookup"><span data-stu-id="17fd9-133">Open `src/App.vue` and replace the file contents with the following code:</span></span>

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

## <a name="start-the-dev-server"></a><span data-ttu-id="17fd9-134">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="17fd9-134">Start the dev server</span></span>

1. <span data-ttu-id="17fd9-135">Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="17fd9-135">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="17fd9-136">В веб-браузере перейдите по адресу `https://localhost:3000` (обратите внимание на `https`).</span><span class="sxs-lookup"><span data-stu-id="17fd9-136">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="17fd9-137">Если появится сообщение, что сертификат сайта не является доверенным, [сделайте так, чтобы компьютер ему доверял](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span><span class="sxs-lookup"><span data-stu-id="17fd9-137">If your browser indicates that the site's certificate is not trusted, you will need to [configure your computer to trust the certificate](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).</span></span>

3. <span data-ttu-id="17fd9-138">Если страница на сайте `https://localhost:3000` пуста, при этом нет ошибок сертификата, значит она работает.</span><span class="sxs-lookup"><span data-stu-id="17fd9-138">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="17fd9-139">Приложение Vue подключается после запуска Office, поэтому в нем отображаются только элементы из среды Excel.</span><span class="sxs-lookup"><span data-stu-id="17fd9-139">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="17fd9-140">Проверка</span><span class="sxs-lookup"><span data-stu-id="17fd9-140">Try it out</span></span>

1. <span data-ttu-id="17fd9-141">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="17fd9-141">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="17fd9-142">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="17fd9-142">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="17fd9-143">Веб-браузер: [загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="17fd9-143">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="17fd9-144">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="17fd9-144">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="17fd9-145">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="17fd9-145">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="17fd9-147">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="17fd9-147">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="17fd9-148">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="17fd9-148">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="17fd9-150">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="17fd9-150">Next steps</span></span>

<span data-ttu-id="17fd9-151">Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Vue!</span><span class="sxs-lookup"><span data-stu-id="17fd9-151">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="17fd9-152">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="17fd9-152">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="17fd9-153">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="17fd9-153">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="17fd9-154">См. также</span><span class="sxs-lookup"><span data-stu-id="17fd9-154">See also</span></span>

* [<span data-ttu-id="17fd9-155">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="17fd9-155">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="17fd9-156">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="17fd9-156">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="17fd9-157">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="17fd9-157">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="17fd9-158">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="17fd9-158">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="17fd9-159">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="17fd9-159">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="17fd9-160">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="17fd9-160">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
