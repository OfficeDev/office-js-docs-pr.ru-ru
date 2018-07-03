# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="e8a74-101">Создание надстройки Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="e8a74-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="e8a74-102">В этой статье описывается процесс создания надстройки Excel с помощью Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="e8a74-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e8a74-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="e8a74-103">Prerequisites</span></span>

- <span data-ttu-id="e8a74-104">Посмотрите, что [необходимо для использования CLI для Angular](https://github.com/angular/angular-cli#prerequisites), и установите все недостающие компоненты.</span><span class="sxs-lookup"><span data-stu-id="e8a74-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="e8a74-105">Глобально установите [CLI для Angular](https://github.com/angular/angular-cli).</span><span class="sxs-lookup"><span data-stu-id="e8a74-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="e8a74-106">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="e8a74-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="e8a74-107">Создание приложения Angular</span><span class="sxs-lookup"><span data-stu-id="e8a74-107">Generate a new Angular app</span></span>

<span data-ttu-id="e8a74-108">Используйте Angular CLI, чтобы создать приложение Angular.</span><span class="sxs-lookup"><span data-stu-id="e8a74-108">Use the Angular CLI to generate your Angular app.</span></span> <span data-ttu-id="e8a74-109">Используя терминал, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="e8a74-109">From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="e8a74-110">Создание файла манифеста</span><span class="sxs-lookup"><span data-stu-id="e8a74-110">Generate the manifest file</span></span>

<span data-ttu-id="e8a74-111">В файле манифеста надстройки определяются ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="e8a74-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="e8a74-112">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="e8a74-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="e8a74-113">Используя генератор Yeoman, создайте файл манифеста для надстройки.</span><span class="sxs-lookup"><span data-stu-id="e8a74-113">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="e8a74-114">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="e8a74-114">Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="e8a74-115">**Выберите тип проекта:** `Manifest`</span><span class="sxs-lookup"><span data-stu-id="e8a74-115">**Choose a project type:** `Manifest`</span></span>
    - <span data-ttu-id="e8a74-116">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="e8a74-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="e8a74-117">**Какое клиентское приложение Office должно поддерживаться?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="e8a74-117">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="e8a74-118">После завершения работы мастера вы сможете создать файл манифеста и файл ресурсов для создания вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="e8a74-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="e8a74-120">Если вам будет предложено переписать файл **package.json**, выберите **No** (Нет).</span><span class="sxs-lookup"><span data-stu-id="e8a74-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="e8a74-121">Защита приложения</span><span class="sxs-lookup"><span data-stu-id="e8a74-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="e8a74-122">Для краткого руководства, вы можете использовать сертификаты, которые предоставляет **генератор Yeoman для надстроек Office**.</span><span class="sxs-lookup"><span data-stu-id="e8a74-122">For this quickstart, you can use the certificates that the **Yeoman generator for Office Add-ins** provides.</span></span> <span data-ttu-id="e8a74-123">Вы уже установили генератор глобально (он входит в список **необходимых компонентов** этого краткого руководства), поэтому вам просто нужно скопировать сертификаты из приложения глобальной установки в папку приложения.</span><span class="sxs-lookup"><span data-stu-id="e8a74-123">You've already installed the generator globally (as part of the **Prerequisites** for this quickstart), so you'll just need to copy the certificates from the global install location into your app folder.</span></span> <span data-ttu-id="e8a74-124">Ниже описано, как выполнить этот процесс.</span><span class="sxs-lookup"><span data-stu-id="e8a74-124">The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="e8a74-125">Используя терминал, выполните следующую команду, чтобы определить папку, в которую установлены глобальные библиотек **npm**:</span><span class="sxs-lookup"><span data-stu-id="e8a74-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="e8a74-126">Первая строка выходных данных, создаваемых этой командой, указывает папку, в которую установлены глобальные библиотеки **npm**.</span><span class="sxs-lookup"><span data-stu-id="e8a74-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="e8a74-127">Используя проводник, перейдите к папке `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`.</span><span class="sxs-lookup"><span data-stu-id="e8a74-127">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder.</span></span> <span data-ttu-id="e8a74-128">Оттуда скопируйте папку `certs` в буфер обмена.</span><span class="sxs-lookup"><span data-stu-id="e8a74-128">From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="e8a74-129">Перейдите в корневую папку приложения Angular, созданную на шаге 1 предыдущего раздела, и вставьте папку `certs` из буфера обмена в эту папку.</span><span class="sxs-lookup"><span data-stu-id="e8a74-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="e8a74-130">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="e8a74-130">Update the app</span></span>

1. <span data-ttu-id="e8a74-131">В редакторе кода откройте файл **package.json** в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="e8a74-131">In your code editor, open **package.json** in the root of the project.</span></span> <span data-ttu-id="e8a74-132">Измените скрипт `start`, чтобы указать, что сервер должен использовать SSL и порт 3000, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-132">Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="e8a74-133">Откройте файл **.angular cli.json** в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="e8a74-133">Open **.angular-cli.json** in the root of the project.</span></span> <span data-ttu-id="e8a74-134">Измените объект **defaults**, чтобы указать расположение сертификатов, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-134">Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. <span data-ttu-id="e8a74-135">Откройте файл **src/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните.</span><span class="sxs-lookup"><span data-stu-id="e8a74-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="e8a74-136">Откройте **src/main.ts**, замените `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="e8a74-137">Откройте **src/polyfills.ts**, добавьте приведенную ниже строку кода над всеми имеющимися операторами `import` и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="e8a74-138">В файле **src/polyfills.ts** раскомментируйте приведенные ниже строки и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. <span data-ttu-id="e8a74-139">Откройте **src/app/app.component.html**, замените его содержимое приведенным ниже кодом HTML и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. <span data-ttu-id="e8a74-140">Откройте **src/app/app.component.css**, замените его содержимое приведенным ниже кодом CSS и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="e8a74-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
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
    ```

9. <span data-ttu-id="e8a74-141">Откройте файл **src/app/app.component.ts**, замените его содержимое приведенным ниже кодом и сохраните.</span><span class="sxs-lookup"><span data-stu-id="e8a74-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="e8a74-142">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="e8a74-142">Start the dev server</span></span>

1. <span data-ttu-id="e8a74-143">Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="e8a74-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="e8a74-p107">В веб-браузере перейдите по адресу `https://localhost:3000`. Если появится сообщение о том, что сертификат сайта не является доверенным, укажите, что ему можно доверять. Дополнительные сведения см. в статье [Добавление самозаверяющего сертификата как доверенного корневого сертификата](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="e8a74-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8a74-147">Chrome может продолжать показывать предупреждение о том, что рабочая станция не доверяет сертификату сайта даже после [его добавления в список доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="e8a74-147">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="e8a74-148">Вы можете игнорировать это предупреждение. Чтобы убедиться, что сертификат является доверенным, перейдите по адресу `https://localhost:3000` в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="e8a74-148">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="e8a74-149">После того как браузер загрузит страницу надстройки без ошибок сертификата, вы можете протестировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="e8a74-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="e8a74-150">Проверка</span><span class="sxs-lookup"><span data-stu-id="e8a74-150">Try it out</span></span>

1. <span data-ttu-id="e8a74-151">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="e8a74-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="e8a74-152">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e8a74-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="e8a74-153">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="e8a74-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="e8a74-154">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="e8a74-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="e8a74-155">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="e8a74-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="e8a74-157">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="e8a74-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="e8a74-158">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="e8a74-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="e8a74-160">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="e8a74-160">Next steps</span></span>

<span data-ttu-id="e8a74-p109">Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="e8a74-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e8a74-163">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="e8a74-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="e8a74-164">См. также</span><span class="sxs-lookup"><span data-stu-id="e8a74-164">See also</span></span>

* [<span data-ttu-id="e8a74-165">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="e8a74-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="e8a74-166">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e8a74-166">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="e8a74-167">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="e8a74-167">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="e8a74-168">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="e8a74-168">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
