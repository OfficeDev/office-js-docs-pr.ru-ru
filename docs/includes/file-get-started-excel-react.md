# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="2603f-101">Создание надстройки Excel с помощью React</span><span class="sxs-lookup"><span data-stu-id="2603f-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="2603f-102">Эта статья ознакомит вас с процессом создания надстройки Excel с помощью React и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2603f-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="2603f-103">Среда</span><span class="sxs-lookup"><span data-stu-id="2603f-103">Environment</span></span>

- <span data-ttu-id="2603f-104">**Классическое приложение Office.** Убедитесь, что у вас установлена ​​последняя версия Office.</span><span class="sxs-lookup"><span data-stu-id="2603f-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="2603f-105">Команды надстроек требуют сборку 16.0.6769.0000 или более позднюю (рекомендуется сборка **16.0.6868.0000**).</span><span class="sxs-lookup"><span data-stu-id="2603f-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="2603f-106">Узнайте, как [установить последнюю версию приложений Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="2603f-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="2603f-107">**Office Online.** Не требуется выполнять дополнительную настройку.</span><span class="sxs-lookup"><span data-stu-id="2603f-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="2603f-108">Обратите внимание, что поддержка команд в Office Online для рабочих и учебных учетных записей предоставляется в тестовом режиме.</span><span class="sxs-lookup"><span data-stu-id="2603f-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2603f-109">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="2603f-109">Prerequisites</span></span>

- <span data-ttu-id="2603f-110">Глобально установите [Create React App](https://github.com/facebookincubator/create-react-app).</span><span class="sxs-lookup"><span data-stu-id="2603f-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="2603f-111">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="2603f-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="2603f-112">Создание приложения React</span><span class="sxs-lookup"><span data-stu-id="2603f-112">Generate a new React app</span></span>

<span data-ttu-id="2603f-113">Создайте приложение React с помощью Create React App.</span><span class="sxs-lookup"><span data-stu-id="2603f-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="2603f-114">В терминале выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="2603f-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="2603f-115">Создание файла манифеста и загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="2603f-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="2603f-116">Каждой надстройке необходим файл манифеста, чтобы определить ее параметры и возможности.</span><span class="sxs-lookup"><span data-stu-id="2603f-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="2603f-117">Перейдите к папке приложения.</span><span class="sxs-lookup"><span data-stu-id="2603f-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="2603f-118">С помощью генератора Yeoman создайте файл манифеста для надстройки.</span><span class="sxs-lookup"><span data-stu-id="2603f-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="2603f-119">Выполните приведенную ниже команду и ответьте на вопросы, как показано на следующем снимке экрана:</span><span class="sxs-lookup"><span data-stu-id="2603f-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="2603f-120">**Выберите тип проекта:** `Manifest`</span><span class="sxs-lookup"><span data-stu-id="2603f-120">**Choose a project type:** `Manifest`</span></span>
    - <span data-ttu-id="2603f-121">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="2603f-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="2603f-122">**Какое клиентское приложение Office должно поддерживаться?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="2603f-122">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="2603f-123">После завершения работы мастера вы сможете создать файл манифеста и файл ресурсов для создания вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="2603f-123">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>
    
    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="2603f-125">Если вам будет предложено переписать файл **package.json**, выберите **No** (не переписывать).</span><span class="sxs-lookup"><span data-stu-id="2603f-125">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="2603f-126">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="2603f-126">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="2603f-127">Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2603f-127">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="2603f-128">Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="2603f-128">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="2603f-129">iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="2603f-129">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="2603f-130">Обновление приложения</span><span class="sxs-lookup"><span data-stu-id="2603f-130">Update the app</span></span>

1. <span data-ttu-id="2603f-131">Откройте **public/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2603f-131">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="2603f-132">Откройте **src/index.js**, замените `ReactDOM.render(<App />, document.getElementById('root'));` приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2603f-132">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="2603f-133">Откройте **src/App.js**, замените его содержимое приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2603f-133">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. <span data-ttu-id="2603f-134">Откройте **src/App.css**, замените его содержимое приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="2603f-134">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

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

## <a name="try-it-out"></a><span data-ttu-id="2603f-135">Проверка</span><span class="sxs-lookup"><span data-stu-id="2603f-135">Try it out</span></span>

1. <span data-ttu-id="2603f-136">Выполните в терминале приведенную ниже команду, чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="2603f-136">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="2603f-137">Windows:</span><span class="sxs-lookup"><span data-stu-id="2603f-137">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="2603f-138">macOS:</span><span class="sxs-lookup"><span data-stu-id="2603f-138">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="2603f-p105">Откроется окно браузера с надстройкой. Закройте это окно.</span><span class="sxs-lookup"><span data-stu-id="2603f-p105">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="2603f-141">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="2603f-141">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="2603f-143">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="2603f-143">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="2603f-144">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="2603f-144">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="2603f-146">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="2603f-146">Next steps</span></span>

<span data-ttu-id="2603f-p106">Поздравляем, вы успешно создали надстройку Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="2603f-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="2603f-149">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="2603f-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="2603f-150">См. также</span><span class="sxs-lookup"><span data-stu-id="2603f-150">See also</span></span>

* [<span data-ttu-id="2603f-151">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="2603f-151">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="2603f-152">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2603f-152">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="2603f-153">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="2603f-153">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="2603f-154">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2603f-154">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
