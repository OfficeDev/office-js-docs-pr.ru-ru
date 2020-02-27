---
ms.date: 02/20/2020
title: Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)
ms.prod: excel
description: Осуществляйте обмен данными и событиями между пользовательскими функциями и областью задач в Excel.
localization_priority: Priority
ms.openlocfilehash: 13ef4c199f7cb1de84e58f0ada554c851aee0cad
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283893"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="b1b9b-103">Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="b1b9b-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="b1b9b-104">Вы можете настроить свою надстройку Excel для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="b1b9b-105">Это позволит предоставлять общий доступ к глобальным данным или отправлять события между областью задач и пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-105">This will make it possible to shared global data, or send events between the task pane and custom functions.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="b1b9b-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="b1b9b-106">Create the add-in project</span></span>

<span data-ttu-id="b1b9b-107">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-107">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="b1b9b-108">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-108">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="b1b9b-109">Выберите тип проекта: **проект надстройки пользовательских функций Excel**</span><span class="sxs-lookup"><span data-stu-id="b1b9b-109">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="b1b9b-110">Выберите тип сценария: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="b1b9b-110">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="b1b9b-111">Как вы хотите назвать надстройку? **Моя надстройка Office**</span><span class="sxs-lookup"><span data-stu-id="b1b9b-111">What do you want to name your add-in? **My Office Add-in**</span></span>

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

<span data-ttu-id="b1b9b-113">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b1b9b-114">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="b1b9b-114">Configure the manifest</span></span>

1. <span data-ttu-id="b1b9b-115">Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-115">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="b1b9b-116">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-116">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="b1b9b-117">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-117">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="b1b9b-118">Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-118">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="b1b9b-119">В элементе `<Page>` измените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-119">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="b1b9b-120">В разделе `<DesktopFormFactor>` измените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-120">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="b1b9b-121">В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-121">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="b1b9b-122">Добавьте новый **Url-идентификатор** для **ContosoAddin.Url**, указывающий на **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-122">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="b1b9b-123">Сохраните изменения и перестройте проект.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-123">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="b1b9b-124">Общий доступ к состоянию для пользовательской функции и кода области задач</span><span class="sxs-lookup"><span data-stu-id="b1b9b-124">Share state between custom function and task pane code</span></span>

<span data-ttu-id="b1b9b-125">Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-125">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="b1b9b-126">В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-126">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="b1b9b-127">Создание пользовательских функций для получения или сохранения общего состояния</span><span class="sxs-lookup"><span data-stu-id="b1b9b-127">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="b1b9b-128">В Visual Studio Code откройте файл **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-128">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="b1b9b-129">В строке 1 в самом верху вставьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-129">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="b1b9b-130">При этом будет инициализирована глобальная переменная **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-130">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="b1b9b-131">Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-131">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="b1b9b-132">Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-132">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="b1b9b-133">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-133">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="b1b9b-134">Создание элементов управления области задач для работы с глобальными данными</span><span class="sxs-lookup"><span data-stu-id="b1b9b-134">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="b1b9b-135">Откройте файл **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-135">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="b1b9b-136">Добавьте следующий элемент скрипта непосредственно перед элементом `</head>`.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-136">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="b1b9b-137">После закрытия элемента `</main>` добавьте следующий HTML-код.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-137">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="b1b9b-138">С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-138">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
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

4. <span data-ttu-id="b1b9b-139">Перед элементом `<body>` добавьте приведенный ниже сценарий.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-139">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="b1b9b-140">Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-140">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="b1b9b-141">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-141">Save the file.</span></span>
6. <span data-ttu-id="b1b9b-142">Построение проекта</span><span class="sxs-lookup"><span data-stu-id="b1b9b-142">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="b1b9b-143">Обмен данными между пользовательскими функциями и областью задач</span><span class="sxs-lookup"><span data-stu-id="b1b9b-143">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="b1b9b-144">Запустите проект, выполнив приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-144">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="b1b9b-145">После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-145">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="b1b9b-146">Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-146">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="b1b9b-147">Можно также использовать `=CONTOSO.STOREVALUE(“new value”)` для изменения значения общих данных.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-147">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="b1b9b-148">Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-148">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="b1b9b-149">Вызов API Office из пользовательских функций не поддерживается в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="b1b9b-149">Calling Office APIs from custom functions is not supported in the preview.</span></span>
