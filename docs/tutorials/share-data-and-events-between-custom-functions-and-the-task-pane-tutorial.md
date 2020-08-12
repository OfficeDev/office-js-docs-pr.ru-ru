---
title: 'Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач'
description: Узнайте, как обмениваться данными и событиями между пользовательскими функциями и областью задач в Excel.
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a48d43270787648d8e5a53c885eab4b69cd8842e
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641153"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="f0580-103">Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач</span><span class="sxs-lookup"><span data-stu-id="f0580-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="f0580-104">Вы можете настроить свою надстройку Excel для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="f0580-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="f0580-105">Это позволяет предоставлять общий доступ к глобальным данным или отправлять события между областью задач и пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="f0580-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="f0580-106">Для большинства пользовательских функций рекомендуется пользоваться общей средой выполнения, если у вас нет особой причины применять пользовательскую функцию без области задач (без пользовательского интерфейса).</span><span class="sxs-lookup"><span data-stu-id="f0580-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="f0580-107">В этом учебном руководстве предполагается, что вы знакомы с использованием генератора Yo Office для создания проектов надстроек.</span><span class="sxs-lookup"><span data-stu-id="f0580-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="f0580-108">Если вы еще этого не сделали, рекомендуется ознакомиться с [руководством по пользовательским функциям в Excel](./excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="f0580-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="f0580-109">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="f0580-109">Create the add-in project</span></span>

<span data-ttu-id="f0580-110">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f0580-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="f0580-111">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="f0580-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="f0580-112">Выберите тип проекта: **проект надстройки пользовательских функций Excel**</span><span class="sxs-lookup"><span data-stu-id="f0580-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="f0580-113">Выберите тип сценария: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="f0580-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="f0580-114">Как вы хотите назвать надстройку? **Моя надстройка Office**</span><span class="sxs-lookup"><span data-stu-id="f0580-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

<span data-ttu-id="f0580-116">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="f0580-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f0580-117">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="f0580-117">Configure the manifest</span></span>

1. <span data-ttu-id="f0580-118">Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="f0580-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="f0580-119">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="f0580-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="f0580-120">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="f0580-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="f0580-121">Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="f0580-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="f0580-122">В элементе `<Page>` измените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="f0580-122">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="f0580-123">В разделе `<DesktopFormFactor>` измените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="f0580-123">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="f0580-124">В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="f0580-124">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="f0580-125">Добавьте новый **Url-идентификатор** для **ContosoAddin.Url**, указывающий на **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f0580-125">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="f0580-126">Сохраните изменения и перестройте проект.</span><span class="sxs-lookup"><span data-stu-id="f0580-126">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="f0580-127">Общий доступ к состоянию для пользовательской функции и кода области задач</span><span class="sxs-lookup"><span data-stu-id="f0580-127">Share state between custom function and task pane code</span></span>

<span data-ttu-id="f0580-128">Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**.</span><span class="sxs-lookup"><span data-stu-id="f0580-128">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="f0580-129">В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.</span><span class="sxs-lookup"><span data-stu-id="f0580-129">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="f0580-130">Создание пользовательских функций для получения или сохранения общего состояния</span><span class="sxs-lookup"><span data-stu-id="f0580-130">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="f0580-131">В Visual Studio Code откройте файл **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="f0580-131">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="f0580-132">В строке 1 в самом верху вставьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="f0580-132">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="f0580-133">При этом будет инициализирована глобальная переменная **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="f0580-133">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="f0580-134">Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="f0580-134">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="f0580-135">Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="f0580-135">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="f0580-136">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="f0580-136">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="f0580-137">Создание элементов управления области задач для работы с глобальными данными</span><span class="sxs-lookup"><span data-stu-id="f0580-137">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="f0580-138">Откройте файл **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f0580-138">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="f0580-139">Добавьте следующий элемент скрипта непосредственно перед элементом `</head>`.</span><span class="sxs-lookup"><span data-stu-id="f0580-139">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="f0580-140">После закрытия элемента `</main>` добавьте следующий HTML-код.</span><span class="sxs-lookup"><span data-stu-id="f0580-140">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="f0580-141">С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.</span><span class="sxs-lookup"><span data-stu-id="f0580-141">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

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

4. <span data-ttu-id="f0580-142">Перед элементом `<body>` добавьте приведенный ниже сценарий.</span><span class="sxs-lookup"><span data-stu-id="f0580-142">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="f0580-143">Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.</span><span class="sxs-lookup"><span data-stu-id="f0580-143">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="f0580-144">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="f0580-144">Save the file.</span></span>
6. <span data-ttu-id="f0580-145">Построение проекта</span><span class="sxs-lookup"><span data-stu-id="f0580-145">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="f0580-146">Обмен данными между пользовательскими функциями и областью задач</span><span class="sxs-lookup"><span data-stu-id="f0580-146">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="f0580-147">Запустите проект, выполнив приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="f0580-147">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="f0580-148">После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных.</span><span class="sxs-lookup"><span data-stu-id="f0580-148">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="f0580-149">Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные.</span><span class="sxs-lookup"><span data-stu-id="f0580-149">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="f0580-150">Можно также использовать `=CONTOSO.STOREVALUE("new value")` для изменения значения общих данных.</span><span class="sxs-lookup"><span data-stu-id="f0580-150">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="f0580-151">Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст.</span><span class="sxs-lookup"><span data-stu-id="f0580-151">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="f0580-152">Вызов некоторых API Office из пользовательских функций невозможен.</span><span class="sxs-lookup"><span data-stu-id="f0580-152">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="f0580-153">Дополнительные сведения см. в статье [Вызов API Microsoft Excel из пользовательской функции](../excel/call-excel-apis-from-custom-function.md).</span><span class="sxs-lookup"><span data-stu-id="f0580-153">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
