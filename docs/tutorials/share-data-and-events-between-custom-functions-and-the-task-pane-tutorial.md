---
ms.date: 11/04/2019
title: Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)
ms.prod: excel
description: Осуществляйте обмен данными и событиями между пользовательскими функциями и областью задач в Excel.
localization_priority: Priority
ms.openlocfilehash: d86b5bb59dd0da51d5b5472288fa802823d658ce
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217360"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="3a3e2-103">Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="3a3e2-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

<span data-ttu-id="3a3e2-104">Пользовательские функции и область задач в Excel совместно используют глобальные данные и могут вызывать функции друг друга.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-104">Excel custom functions and the task pane share global data, and can make function calls into each other.</span></span> <span data-ttu-id="3a3e2-105">Следуя инструкциям, приведенным в этой статье, настройте проект таким образом, чтобы пользовательские функции могли работать с областью задач.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-105">To configure your project so that custom functions can work with the task pane, follow the instructions in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="3a3e2-106">Возможности, описанные в этой статье, в настоящее время доступны в предварительной версии и могут изменяться.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-106">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="3a3e2-107">Сейчас они не поддерживаются для использования в рабочих средах.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-107">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="3a3e2-108">Возможности предварительной версии, приведенные в этой статье, доступны только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-108">The preview features in this article are only available on Excel on Windows.</span></span> <span data-ttu-id="3a3e2-109">Чтобы ознакомиться с ними, вам нужно [присоединиться к программе предварительной оценки Office](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="3a3e2-109">To try the preview features, you will need to [join Office Insider](https://insider.office.com/join).</span></span>  <span data-ttu-id="3a3e2-110">Хороший способ ознакомиться с такими возможностями — использование подписки на Office 365.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-110">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="3a3e2-111">Если у вас еще нет подписки на Office 365, вы можете оформить бесплатную возобновляемую подписку на Office 365 на 90 дней, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="3a3e2-111">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="3a3e2-112">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="3a3e2-112">Create the add-in project</span></span>

<span data-ttu-id="3a3e2-113">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-113">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="3a3e2-114">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-114">Run the following command and then answer the prompts with the following answers:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="3a3e2-115">Выберите тип проекта: **проект надстройки пользовательских функций Excel**</span><span class="sxs-lookup"><span data-stu-id="3a3e2-115">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="3a3e2-116">Выберите тип сценария: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="3a3e2-116">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="3a3e2-117">Как вы хотите назвать надстройку? **Моя надстройка Office**</span><span class="sxs-lookup"><span data-stu-id="3a3e2-117">What do you want to name your add-in? **My Office Add-in**</span></span>

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

<span data-ttu-id="3a3e2-119">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3a3e2-120">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="3a3e2-120">Configure the manifest</span></span>

1. <span data-ttu-id="3a3e2-121">Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-121">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="3a3e2-122">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-122">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="3a3e2-123">Измените раздел `<Requirements>`, чтобы использовать **CustomFunctionsRuntime** версии **1.2**, как показано в приведенном ниже примере кода.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-123">Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.</span></span>
    
    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. <span data-ttu-id="3a3e2-124">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-124">Find the  `<VersionOverrides>` section and add the following example entry to the `<Runtimes>` section:</span></span> <span data-ttu-id="3a3e2-125">Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-125">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>
    
    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
      <Hosts>
        <Host xsi:type="Workbook">
        <Runtimes>
          <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
        </Runtimes>
        <AllFormFactors>
    ```
    
5. <span data-ttu-id="3a3e2-126">В элементе `<Page>` измените расположение источника с **Functions.Page.Url** на **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. <span data-ttu-id="3a3e2-127">В разделе `<DesktopFormFactor>` измените расположение **FunctionFile** с **Commands.Url** на **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-127">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. <span data-ttu-id="3a3e2-128">В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-128">In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. <span data-ttu-id="3a3e2-129">Добавьте новый **Url-идентификатор** для **TaskPaneAndCustomFunction.Url**, указывающий на **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-129">Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.</span></span>
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. <span data-ttu-id="3a3e2-130">Сохраните изменения и перестройте проект.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-130">Save your changes and rebuild the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="3a3e2-131">Общий доступ к состоянию для пользовательской функции и кода области задач</span><span class="sxs-lookup"><span data-stu-id="3a3e2-131">Share state between custom function and task pane code</span></span> 

<span data-ttu-id="3a3e2-132">Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-132">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="3a3e2-133">В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-133">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="3a3e2-134">Создание пользовательских функций для получения или сохранения общего состояния</span><span class="sxs-lookup"><span data-stu-id="3a3e2-134">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="3a3e2-135">В Visual Studio Code откройте файл **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-135">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="3a3e2-136">В строке 1 в самом верху вставьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-136">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="3a3e2-137">При этом будет инициализирована глобальная переменная **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-137">This will initialize a global variable named **sharedState**.</span></span>
    
    ```js
    window.sharedState = "empty";
    ```
    
3. <span data-ttu-id="3a3e2-138">Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-138">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>
    
    ```js
    /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
    function storeValue(sharedValue) {
    window.sharedState = sharedValue;
    return "value stored";
    }
    ```
    
4. <span data-ttu-id="3a3e2-139">Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-139">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

    ```js
    /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
    function getValue() {
    return window.sharedState;
    }
    ```
    
5. <span data-ttu-id="3a3e2-140">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-140">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="3a3e2-141">Создание элементов управления области задач для работы с глобальными данными</span><span class="sxs-lookup"><span data-stu-id="3a3e2-141">Create task pane controls to work with global data</span></span> 

1. <span data-ttu-id="3a3e2-142">Откройте файл **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-142">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="3a3e2-143">Добавьте следующий элемент скрипта непосредственно перед элементом `</head>`.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-143">Add the following script element just before the `</head>` element.</span></span>

    ```html
    <script src="functions.js"></script>
    ```

3. <span data-ttu-id="3a3e2-144">После закрытия элемента `</main>` добавьте следующий HTML-код.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-144">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="3a3e2-145">С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-145">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
    <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>
    <p>Store new value to shared state</p>
    <div>
    <input type="text" id="storeBox" />
    <button onclick="storeSharedValue()">Store</button>
    </div>
     
    <p>Get shared state value</p>
    <div>
    <input type="text" id="getBox" />
    <button onclick="getSharedValue()">Get</button>
    </div>
    ```
    
4. <span data-ttu-id="3a3e2-146">Перед элементом `<body>` добавьте приведенный ниже сценарий.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-146">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="3a3e2-147">Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-147">This code will handle the button click events when the user wants to store or get global data.</span></span>
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
5. <span data-ttu-id="3a3e2-148">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-148">Save the file.</span></span>
6. <span data-ttu-id="3a3e2-149">Построение проекта</span><span class="sxs-lookup"><span data-stu-id="3a3e2-149">Build the project</span></span>
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="3a3e2-150">Обмен данными между пользовательскими функциями и областью задач</span><span class="sxs-lookup"><span data-stu-id="3a3e2-150">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="3a3e2-151">Запустите проект, выполнив приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-151">Start the project by using the following command.</span></span>

    ```command&nbsp;line
    npm run start
    ```

<span data-ttu-id="3a3e2-152">После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-152">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="3a3e2-153">Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-153">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="3a3e2-154">Можно также использовать `=CONTOSO.STOREVALUE(“new value”)` для изменения значения общих данных.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-154">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="3a3e2-155">Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-155">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="3a3e2-156">Вызов API Office из пользовательских функций не поддерживается в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="3a3e2-156">Calling Office APIs from custom functions is not supported in the preview.</span></span>

