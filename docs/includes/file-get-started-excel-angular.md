# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="93711-101">Создание надстройки Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="93711-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="93711-102">В этой статье описывается процесс создания надстройки Excel с помощью Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="93711-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="93711-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="93711-103">Prerequisites</span></span>

- <span data-ttu-id="93711-104">Посмотрите, что [необходимо для использования CLI для Angular](https://github.com/angular/angular-cli#prerequisites), и установите все недостающие компоненты.</span><span class="sxs-lookup"><span data-stu-id="93711-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="93711-105">Глобально установите [CLI для Angular](https://github.com/angular/angular-cli).</span><span class="sxs-lookup"><span data-stu-id="93711-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="93711-106">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="93711-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="93711-107">Создание приложения Angular</span><span class="sxs-lookup"><span data-stu-id="93711-107">Generate a new Angular app</span></span>

<span data-ttu-id="93711-108">Используйте Angular CLI, чтобы создать приложение Angular.</span><span class="sxs-lookup"><span data-stu-id="93711-108">Use the Angular CLI to generate your Angular app.</span></span> <span data-ttu-id="93711-109">Используя терминал, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="93711-109">From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="93711-110">Создание файла манифеста</span><span class="sxs-lookup"><span data-stu-id="93711-110">Generate the manifest file</span></span>

<span data-ttu-id="93711-111">В файле манифеста надстройки определяются ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="93711-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="93711-112">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="93711-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="93711-113">Используя генератор Yeoman, создайте файл манифеста для надстройки.</span><span class="sxs-lookup"><span data-stu-id="93711-113">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="93711-114">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="93711-114">Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office
    ```
    - <span data-ttu-id="93711-115">**Для **Would you like to create a new subfolder for your project?** (Создать новую вложенную папку для проекта?) выберите `No` (Нет).** `No`</span><span class="sxs-lookup"><span data-stu-id="93711-115">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="93711-116">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="93711-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="93711-117">**Какое клиентское приложение Office должно поддерживаться?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="93711-117">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="93711-118">****Would you like to create a new add-in?:** `No` (Создать новую надстройку?)** `No`</span><span class="sxs-lookup"><span data-stu-id="93711-118">**Would you like to create a new add-in?:** `No`</span></span>

    <span data-ttu-id="93711-p103">Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока генератор закончит работу.</span><span class="sxs-lookup"><span data-stu-id="93711-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="93711-123">Если вам будет предложено переписать файл **package.json**, выберите **No** (Нет).</span><span class="sxs-lookup"><span data-stu-id="93711-123">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="93711-124">Защита приложения</span><span class="sxs-lookup"><span data-stu-id="93711-124">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="93711-125">Для целей этого краткого руководства можно использовать сертификаты, которые предоставляет **генератор Yeoman для надстроек Office**.</span><span class="sxs-lookup"><span data-stu-id="93711-125">For this quickstart, you can use the certificates that the **Yeoman generator for Office Add-ins** provides.</span></span> <span data-ttu-id="93711-126">Вы уже установили генератор глобально (он входит в список **необходимых компонентов** этого краткого руководства), поэтому вам просто нужно скопировать сертификаты из места глобальной установки в папку приложения.</span><span class="sxs-lookup"><span data-stu-id="93711-126">You've already installed the generator globally (as part of the **Prerequisites** for this quickstart), so you'll just need to copy the certificates from the global install location into your app folder.</span></span> <span data-ttu-id="93711-127">Ниже описано, как это сделать.</span><span class="sxs-lookup"><span data-stu-id="93711-127">The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="93711-128">Используя терминал, выполните следующую команду, чтобы определить папку, в которую установлены глобальные библиотек **npm**:</span><span class="sxs-lookup"><span data-stu-id="93711-128">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="93711-129">Первая строка выходных данных, создаваемых этой командой, указывает папку, в которую установлены глобальные библиотеки **npm**.</span><span class="sxs-lookup"><span data-stu-id="93711-129">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="93711-130">Используя проводник, перейдите к папке `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`.</span><span class="sxs-lookup"><span data-stu-id="93711-130">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder.</span></span> <span data-ttu-id="93711-131">Оттуда скопируйте папку `certs` в буфер обмена.</span><span class="sxs-lookup"><span data-stu-id="93711-131">From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="93711-132">Перейдите в корневую папку приложения Angular, созданную на шаге 1 предыдущего раздела, и вставьте папку `certs` из буфера обмена в эту папку.</span><span class="sxs-lookup"><span data-stu-id="93711-132">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="93711-133">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="93711-133">Update the app</span></span>

1. <span data-ttu-id="93711-134">В редакторе кода откройте файл **package.json** в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="93711-134">In your code editor, open **package.json** in the root of the project.</span></span> <span data-ttu-id="93711-135">Измените скрипт `start`, чтобы указать, что сервер должен использовать SSL и порт 3000, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-135">Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="93711-136">Откройте файл **.angular cli.json** в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="93711-136">Open **.angular-cli.json** in the root of the project.</span></span> <span data-ttu-id="93711-137">Измените объект **defaults**, чтобы указать расположение сертификатов, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-137">Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

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

3. <span data-ttu-id="93711-138">Откройте файл **src/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните.</span><span class="sxs-lookup"><span data-stu-id="93711-138">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="93711-139">Откройте **src/main.ts**, замените `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-139">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="93711-140">Откройте **src/polyfills.ts**, добавьте приведенную ниже строку кода над всеми имеющимися операторами `import` и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-140">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="93711-141">В файле **src/polyfills.ts** раскомментируйте приведенные ниже строки и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-141">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

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

7. <span data-ttu-id="93711-142">Откройте **src/app/app.component.html**, замените его содержимое приведенным ниже кодом HTML и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-142">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

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

8. <span data-ttu-id="93711-143">Откройте **src/app/app.component.css**, замените его содержимое приведенным ниже кодом CSS и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="93711-143">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="93711-144">Откройте файл **src/app/app.component.ts**, замените его содержимое приведенным ниже кодом и сохраните.</span><span class="sxs-lookup"><span data-stu-id="93711-144">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

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

## <a name="start-the-dev-server"></a><span data-ttu-id="93711-145">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="93711-145">Start the dev server</span></span>

1. <span data-ttu-id="93711-146">Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="93711-146">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="93711-p108">В веб-браузере перейдите по адресу `https://localhost:3000`. Если появится сообщение о том, что сертификат сайта не является доверенным, укажите, что ему можно доверять. Дополнительные сведения см. в статье [Добавление самозаверяющего сертификата как доверенного корневого сертификата](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="93711-p108">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="93711-150">Chrome может продолжать показывать предупреждение о том, что рабочая станция не доверяет сертификату сайта даже после [его добавления в список доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="93711-150">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="93711-151">Вы можете игнорировать это предупреждение. Чтобы убедиться, что сертификат является доверенным, перейдите по адресу `https://localhost:3000` в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="93711-151">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="93711-152">После того как браузер загрузит страницу надстройки без ошибок сертификата, вы можете протестировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="93711-152">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="93711-153">Проверка</span><span class="sxs-lookup"><span data-stu-id="93711-153">Try it out</span></span>

1. <span data-ttu-id="93711-154">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="93711-154">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="93711-155">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="93711-155">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="93711-156">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="93711-156">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="93711-157">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="93711-157">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="93711-158">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="93711-158">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="93711-160">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="93711-160">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="93711-161">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="93711-161">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="93711-163">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="93711-163">Next steps</span></span>

<span data-ttu-id="93711-p110">Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="93711-p110">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="93711-166">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="93711-166">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="93711-167">См. также</span><span class="sxs-lookup"><span data-stu-id="93711-167">See also</span></span>

* [<span data-ttu-id="93711-168">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="93711-168">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="93711-169">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="93711-169">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="93711-170">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="93711-170">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="93711-171">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="93711-171">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
