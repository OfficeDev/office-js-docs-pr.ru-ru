---
title: 'Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач'
description: Узнайте, как обмениваться данными и событиями между пользовательскими функциями и областью задач в Excel.
ms.date: 08/13/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e4dfb8afc57dc9590d47d927d1f540431d9c8838
ms.sourcegitcommit: 3efa932b70035dde922929d207896e1a6007f620
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/15/2020
ms.locfileid: "46757382"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="4b694-103">Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач</span><span class="sxs-lookup"><span data-stu-id="4b694-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="4b694-104">Вы можете настроить свою надстройку Excel для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="4b694-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="4b694-105">Это позволяет предоставлять общий доступ к глобальным данным или отправлять события между областью задач и пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="4b694-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="4b694-106">Для большинства пользовательских функций рекомендуется пользоваться общей средой выполнения, если у вас нет особой причины применять пользовательскую функцию без области задач (без пользовательского интерфейса).</span><span class="sxs-lookup"><span data-stu-id="4b694-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="4b694-107">В этом учебном руководстве предполагается, что вы знакомы с использованием генератора Yo Office для создания проектов надстроек.</span><span class="sxs-lookup"><span data-stu-id="4b694-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="4b694-108">Если вы еще этого не сделали, рекомендуется ознакомиться с [руководством по пользовательским функциям в Excel](./excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="4b694-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="4b694-109">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="4b694-109">Create the add-in project</span></span>

<span data-ttu-id="4b694-110">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="4b694-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="4b694-111">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="4b694-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="4b694-112">Выберите тип проекта: **проект надстройки пользовательских функций Excel**</span><span class="sxs-lookup"><span data-stu-id="4b694-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="4b694-113">Выберите тип сценария: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="4b694-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="4b694-114">Как вы хотите назвать надстройку? **Моя надстройка Office**</span><span class="sxs-lookup"><span data-stu-id="4b694-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

<span data-ttu-id="4b694-116">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="4b694-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="4b694-117">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="4b694-117">Configure the manifest</span></span>

1. <span data-ttu-id="4b694-118">Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="4b694-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="4b694-119">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="4b694-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="4b694-120">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="4b694-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="4b694-121">Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="4b694-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

> [!NOTE]
> <span data-ttu-id="4b694-122">Если в манифесте надстройки есть элемент `Runtimes`, она использует Internet Explorer 11 независимо от того, какая у вас версия Windows или Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="4b694-122">If your add-in includes the `Runtimes` element in the manifest, it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="4b694-123">Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).</span><span class="sxs-lookup"><span data-stu-id="4b694-123">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

4. <span data-ttu-id="4b694-124">В элементе `<Page>` замените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="4b694-124">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="4b694-125">В разделе `<DesktopFormFactor>` измените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="4b694-125">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="4b694-126">В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="4b694-126">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="4b694-127">Добавьте новый **Url-идентификатор** для **ContosoAddin.Url**, указывающий на **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="4b694-127">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="4b694-128">Сохраните изменения и перестройте проект.</span><span class="sxs-lookup"><span data-stu-id="4b694-128">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="4b694-129">Общий доступ к состоянию для пользовательской функции и кода области задач</span><span class="sxs-lookup"><span data-stu-id="4b694-129">Share state between custom function and task pane code</span></span>

<span data-ttu-id="4b694-130">Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**.</span><span class="sxs-lookup"><span data-stu-id="4b694-130">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="4b694-131">В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.</span><span class="sxs-lookup"><span data-stu-id="4b694-131">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="4b694-132">Создание пользовательских функций для получения или сохранения общего состояния</span><span class="sxs-lookup"><span data-stu-id="4b694-132">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="4b694-133">В Visual Studio Code откройте файл **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="4b694-133">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="4b694-134">В строке 1 в самом верху вставьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="4b694-134">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="4b694-135">При этом будет инициализирована глобальная переменная **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="4b694-135">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="4b694-136">Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="4b694-136">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="4b694-137">Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="4b694-137">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="4b694-138">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="4b694-138">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="4b694-139">Создание элементов управления области задач для работы с глобальными данными</span><span class="sxs-lookup"><span data-stu-id="4b694-139">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="4b694-140">Откройте файл **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="4b694-140">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="4b694-141">Добавьте следующий элемент скрипта непосредственно перед элементом `</head>`.</span><span class="sxs-lookup"><span data-stu-id="4b694-141">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="4b694-142">После закрытия элемента `</main>` добавьте следующий HTML-код.</span><span class="sxs-lookup"><span data-stu-id="4b694-142">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="4b694-143">С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.</span><span class="sxs-lookup"><span data-stu-id="4b694-143">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

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

4. <span data-ttu-id="4b694-144">Перед элементом `<body>` добавьте приведенный ниже сценарий.</span><span class="sxs-lookup"><span data-stu-id="4b694-144">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="4b694-145">Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.</span><span class="sxs-lookup"><span data-stu-id="4b694-145">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="4b694-146">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="4b694-146">Save the file.</span></span>
6. <span data-ttu-id="4b694-147">Построение проекта</span><span class="sxs-lookup"><span data-stu-id="4b694-147">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="4b694-148">Обмен данными между пользовательскими функциями и областью задач</span><span class="sxs-lookup"><span data-stu-id="4b694-148">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="4b694-149">Запустите проект, выполнив приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="4b694-149">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="4b694-150">После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных.</span><span class="sxs-lookup"><span data-stu-id="4b694-150">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="4b694-151">Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные.</span><span class="sxs-lookup"><span data-stu-id="4b694-151">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="4b694-152">Можно также использовать `=CONTOSO.STOREVALUE("new value")` для изменения значения общих данных.</span><span class="sxs-lookup"><span data-stu-id="4b694-152">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="4b694-153">Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст.</span><span class="sxs-lookup"><span data-stu-id="4b694-153">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="4b694-154">Вызов некоторых API Office из пользовательских функций невозможен.</span><span class="sxs-lookup"><span data-stu-id="4b694-154">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="4b694-155">Дополнительные сведения см. в статье [Вызов API Microsoft Excel из пользовательской функции](../excel/call-excel-apis-from-custom-function.md).</span><span class="sxs-lookup"><span data-stu-id="4b694-155">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
