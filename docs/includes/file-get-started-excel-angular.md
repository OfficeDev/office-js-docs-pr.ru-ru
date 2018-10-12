# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="dfaaf-101">Создание надстройки Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="dfaaf-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="dfaaf-102">В этой статье описывается процесс создания надстройки Excel с помощью Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dfaaf-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="dfaaf-103">Prerequisites</span></span>

- <span data-ttu-id="dfaaf-104">Посмотрите, что [необходимо для использования CLI для Angular](https://github.com/angular/angular-cli#prerequisites), и установите все недостающие компоненты.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="dfaaf-105">Глобально установите [CLI для Angular](https://github.com/angular/angular-cli).</span><span class="sxs-lookup"><span data-stu-id="dfaaf-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="dfaaf-106">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="dfaaf-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="dfaaf-107">Создание приложения Angular</span><span class="sxs-lookup"><span data-stu-id="dfaaf-107">Generate a new Angular app</span></span>

<span data-ttu-id="dfaaf-p101">Используйте угловые CLI, чтобы создать угловые приложения. Используя терминал, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p101">Use the Angular CLI to generate your Angular app. From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="dfaaf-110">Создание файла манифеста</span><span class="sxs-lookup"><span data-stu-id="dfaaf-110">Generate the manifest file</span></span>

<span data-ttu-id="dfaaf-111">В файле манифеста надстройки определяются ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="dfaaf-112">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="dfaaf-p102">Используя генератор Yeoman, создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на запросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p102">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="dfaaf-115">**Выберите тип проекта:** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="dfaaf-115">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="dfaaf-116">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="dfaaf-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="dfaaf-117">**Какое клиентское приложение Office вы хотели бы поддерживать?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="dfaaf-117">**Which Office client application would you like to support?:** `Excel`</span></span>

    <span data-ttu-id="dfaaf-118">После завершения работы мастера вы сможете создать файл манифеста и файл ресурсов для создания вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="dfaaf-120">Если вам будет предложено перезаписать файл **package.json**, выберите **No** (Не перезаписывать).</span><span class="sxs-lookup"><span data-stu-id="dfaaf-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="dfaaf-121">Защита приложения</span><span class="sxs-lookup"><span data-stu-id="dfaaf-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="dfaaf-p103">Для краткого руководства можно использовать сертификаты, которые предоставляют **Генератор Yeoman для надстроек Office**. Вы уже установили генератор глобально  (как часть **Необходимых компонентов** этого краткого руководства), поэтому вам просто нужно скопировать сертификаты из приложения глобальной установки в папку приложения. Следуюшие шаги описывают как выполнить этот процесс.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p103">For this quick start, you can use the certificates that the **Yeoman generator for Office Add-ins** provides. You've already installed the generator globally (as part of the **Prerequisites** for this quick start), so you'll just need to copy the certificates from the global install location into your app folder. The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="dfaaf-125">Используя терминал, выполните следующую команду, чтобы определить папку, в которую установлены глобальные библиотеки **npm**:</span><span class="sxs-lookup"><span data-stu-id="dfaaf-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="dfaaf-126">Первая строка выходных данных, создаваемых этой командой, указывает папку, в которую установлены глобальные библиотеки **npm**.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="dfaaf-p104">Используя проводник, перейдите к папке  `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` . Из этого расположения скопируйте папку `certs` в буфер обмена.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p104">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder. From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="dfaaf-129">Перейдите в корневую папку приложения Angular, созданную на шаге 1 предыдущего раздела, и вставьте папку `certs` из буфера обмена в эту папку.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="dfaaf-130">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="dfaaf-130">Update the app</span></span>

1. <span data-ttu-id="dfaaf-p105">В редакторе кода откройте файл **package.json** в корневой папке проекта. Измените скрипт `start`, чтобы указать, что сервер должен использовать SSL и порт 3000, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p105">In your code editor, open **package.json** in the root of the project. Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="dfaaf-p106">Откройте файл **.angular cli.json** в корневой папке проекта. Измените объект **defaults**, чтобы указать расположение сертификатов, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p106">Open **.angular-cli.json** in the root of the project. Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

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

3. <span data-ttu-id="dfaaf-135">Откройте файл **src/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="dfaaf-136">Откройте файл **src/main.ts**, замените `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="dfaaf-137">Откройте файл **src/polyfills.ts**, добавьте приведенную ниже строку кода над всеми имеющимися операторами `import` и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="dfaaf-138">В файле **src/polyfills.ts** раскомментируйте приведенные ниже строки и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

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

7. <span data-ttu-id="dfaaf-139">Откройте файл **src/app/app.component.html**, замените его содержимое приведенным ниже кодом HTML и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

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

8. <span data-ttu-id="dfaaf-140">Откройте файл **src/app/app.component.css**, замените его содержимое приведенным ниже кодом CSS и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="dfaaf-141">Откройте файл **src/app/app.component.ts**, замените его содержимое приведенным ниже кодом и сохраните.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

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

## <a name="start-the-dev-server"></a><span data-ttu-id="dfaaf-142">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="dfaaf-142">Start the dev server</span></span>

1. <span data-ttu-id="dfaaf-143">Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="dfaaf-p107">В веб-браузере перейдите по адресу `https://localhost:3000`. Если появится сообщение о том, что сертификат сайта не является доверенным, укажите, что ему можно доверять. Дополнительные сведения см. в статье [Добавление самозаверяющего сертификата как доверенного корневого сертификата](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="dfaaf-p108">Веб-браузер Chrome может продолжать показывать предупреждение о том, что рабочая станция не доверяет сертификату сайта даже после его [добавления в список доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Вы можете игнорировать это предупреждение. Чтобы убедиться, что сертификат является доверенным, перейдите по адресу  `https://localhost:3000` в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p108">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="dfaaf-149">После того как браузер загрузит страницу надстройки без ошибок сертификата, вы можете протестировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="dfaaf-150">Проверка</span><span class="sxs-lookup"><span data-stu-id="dfaaf-150">Try it out</span></span>

1. <span data-ttu-id="dfaaf-151">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="dfaaf-152">Windows: [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="dfaaf-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="dfaaf-153">Excel Online: [загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="dfaaf-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="dfaaf-154">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="dfaaf-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="dfaaf-155">В Excel перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="dfaaf-157">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="dfaaf-158">В области задач нажмите кнопку **Задать цвет**, чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="dfaaf-160">Дальнейшие шаги</span><span class="sxs-lookup"><span data-stu-id="dfaaf-160">Next steps</span></span>

<span data-ttu-id="dfaaf-p109">Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="dfaaf-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="dfaaf-163">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="dfaaf-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="dfaaf-164">См. также</span><span class="sxs-lookup"><span data-stu-id="dfaaf-164">See also</span></span>

* [<span data-ttu-id="dfaaf-165">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="dfaaf-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="dfaaf-166">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="dfaaf-166">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="dfaaf-167">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="dfaaf-167">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="dfaaf-168">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="dfaaf-168">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
