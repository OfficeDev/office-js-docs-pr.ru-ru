<span data-ttu-id="fe540-101">На данном этапе, описанном в руководстве, вы добавите на ленту еще одну кнопку, при нажатии которой будет выполнена определенная вами функция включения или выключения защиты листа.</span><span class="sxs-lookup"><span data-stu-id="fe540-101">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

> [!NOTE]
> <span data-ttu-id="fe540-102">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="fe540-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="fe540-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="fe540-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="fe540-104">Настройка манифеста для добавления второй кнопки на ленту</span><span class="sxs-lookup"><span data-stu-id="fe540-104">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="fe540-105">Откройте файл манифеста **my-office-add-in-manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="fe540-105">Open the manifest file **my-office-add-in-manifest.xml**.</span></span>
2. <span data-ttu-id="fe540-106">Найдите элемент `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="fe540-106">Find the `<Control>` element.</span></span> <span data-ttu-id="fe540-107">Этот элемент определяет кнопку **Show Taskpane** (Показать область задач) на вкладке **Главная**, которую вы используете для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="fe540-107">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="fe540-108">Мы добавим вторую кнопку в эту же группу на ленте **Главная**.</span><span class="sxs-lookup"><span data-stu-id="fe540-108">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="fe540-109">Добавьте приведенный ниже код между закрывающим тегом элемента управления (`</Control>`) и закрывающим тегом группы (`</Group>`).</span><span class="sxs-lookup"><span data-stu-id="fe540-109">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="fe540-110">Замените `TODO1` строкой, которая присваивает кнопке идентификатор, уникальный в пределах этого файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="fe540-110">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="fe540-111">В манифесте есть только еще одна кнопка, поэтому выполнить задачу несложно.</span><span class="sxs-lookup"><span data-stu-id="fe540-111">There's only one other button in the manifest, so this isn't difficult.</span></span> <span data-ttu-id="fe540-112">Так как кнопка будет включать и выключать защиту листа, укажите "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="fe540-112">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="fe540-113">Когда сделаете это, весь открывающий тег элемента управления должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fe540-113">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="fe540-114">Следующие три элемента `TODO` устанавливают "resid", или идентификаторы ресурса.</span><span class="sxs-lookup"><span data-stu-id="fe540-114">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="fe540-115">Ресурс должен быть строкой, и вы создадите эти три строки на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="fe540-115">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="fe540-116">Сейчас вам нужно присвоить идентификаторы ресурсам.</span><span class="sxs-lookup"><span data-stu-id="fe540-116">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="fe540-117">Кнопка должна называться "Toggle Protection" (Переключение защиты), но у строки должен быть *идентификатор* "ProtectionButtonLabel", поэтому готовый элемент `Label` выглядит следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fe540-117">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="fe540-118">Элемент `SuperTip` определяет подсказку для кнопки.</span><span class="sxs-lookup"><span data-stu-id="fe540-118">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="fe540-119">Заголовок этой подсказки должен совпадать с названием кнопки, поэтому мы используем тот же ИД ресурса — "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="fe540-119">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="fe540-120">Описание подсказки будет следующим: "Click to turn protection of the worksheet on and off" (Нажмите для включения или выключения защиты листа).</span><span class="sxs-lookup"><span data-stu-id="fe540-120">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="fe540-121">У `ID` должно быть значение "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="fe540-121">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="fe540-122">После выполнения весь код `SuperTip` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fe540-122">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="fe540-123">В рабочей надстройке не нужно использовать один и тот же значок для двух разных кнопок, но сейчас мы предлагаем сделать это для простоты.</span><span class="sxs-lookup"><span data-stu-id="fe540-123">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="fe540-124">Поэтому код `Icon` в новом теге `Control` представляет собой лишь копию элемента `Icon` из существующего тега `Control`.</span><span class="sxs-lookup"><span data-stu-id="fe540-124">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="fe540-125">Для элемента `Action` в исходном элементе `Control`, уже присутствующем в манифесте, задан тип `ShowTaskpane`, но новая кнопка будет не открывать область задач, а выполнять специальную функцию, которую вы создадите позже.</span><span class="sxs-lookup"><span data-stu-id="fe540-125">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="fe540-126">Поэтому замените `TODO5` на `ExecuteFunction` (тип действия для кнопок, запускающих специальные функции).</span><span class="sxs-lookup"><span data-stu-id="fe540-126">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="fe540-127">Открывающий тег `Action` должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fe540-127">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="fe540-128">У исходного элемента `Action` есть дочерние элементы, определяющие идентификатор области задач и URL-адрес страницы, которая должна быть открыта в области задач.</span><span class="sxs-lookup"><span data-stu-id="fe540-128">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="fe540-129">Но у элемента `Action` типа `ExecuteFunction` есть один дочерний элемент, который именует функцию, выполняемую элементом управления.</span><span class="sxs-lookup"><span data-stu-id="fe540-129">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="fe540-130">На более позднем этапе вы создадите функцию `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="fe540-130">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="fe540-131">Поэтому замените `TODO6` следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="fe540-131">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="fe540-132">Теперь весь код `Control` должен выглядеть вот так:</span><span class="sxs-lookup"><span data-stu-id="fe540-132">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="fe540-133">Прокрутите страницу вниз до раздела `Resources` манифеста.</span><span class="sxs-lookup"><span data-stu-id="fe540-133">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="fe540-134">Добавьте приведенный ниже код в качестве дочернего элемента `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="fe540-134">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="fe540-135">Добавьте приведенный ниже код в качестве дочернего элемента `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="fe540-135">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="fe540-136">Обязательно сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="fe540-136">Be sure to save the file.</span></span>

## <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="fe540-137">Создание функции защиты листа</span><span class="sxs-lookup"><span data-stu-id="fe540-137">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="fe540-138">Откройте файл \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="fe540-138">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="fe540-139">В файле уже есть функция-выражение, вызываемая сразу после создания (IIFE).</span><span class="sxs-lookup"><span data-stu-id="fe540-139">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="fe540-140">Пользовательская логика инициализации не требуется, поэтому оставьте тело функции, назначенной `Office.initialize`, пустым.</span><span class="sxs-lookup"><span data-stu-id="fe540-140">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="fe540-141">(Но не удаляйте его.</span><span class="sxs-lookup"><span data-stu-id="fe540-141">(But do not delete it.</span></span> <span data-ttu-id="fe540-142">Свойство `Office.initialize` не может быть неопределенным или иметь значение NULL.) *За пределами IIFE* добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="fe540-142">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="fe540-143">Обратите внимание на то, что мы указываем параметр `args` для метода, а самая последняя строка метода вызывает `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="fe540-143">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="fe540-144">Это требование для всех команд надстройки типа **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="fe540-144">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="fe540-145">Это сигнализирует ведущему приложению Office о том, что работа функции завершена и пользовательский интерфейс снова может реагировать.</span><span class="sxs-lookup"><span data-stu-id="fe540-145">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="fe540-146">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="fe540-146">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="fe540-147">В этом коде используется свойство защиты объекта листа в стандартном шаблоне переключателя.</span><span class="sxs-lookup"><span data-stu-id="fe540-147">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="fe540-148">Объяснение `TODO2` будет приведено в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="fe540-148">The `TODO2` will be explained in the next section.</span></span>

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="fe540-149">Добавление кода для получения свойств документа в объекты скрипта области задач</span><span class="sxs-lookup"><span data-stu-id="fe540-149">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="fe540-150">В случае всех описанных ранее функций из этой серии руководств вы ставили в очередь команды для *записи* данных в документ Office.</span><span class="sxs-lookup"><span data-stu-id="fe540-150">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="fe540-151">Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет выставленные в очередь команды документу для выполнения.</span><span class="sxs-lookup"><span data-stu-id="fe540-151">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="fe540-152">Но код, который вы добавили на последнем этапе, вызывает свойство `sheet.protection.protected`, и в этом заключается существенное отличие от ранее написанных функций, так как `sheet` является лишь объектом прокси, существующим в скрипте вашей области задач.</span><span class="sxs-lookup"><span data-stu-id="fe540-152">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="fe540-153">В нем нет сведений о фактическом состоянии защиты документа, поэтому его свойство `protection.protected` не может иметь реального значения.</span><span class="sxs-lookup"><span data-stu-id="fe540-153">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="fe540-154">Сначала нужно получить сведения о состоянии защиты от документа и задать значение `sheet.protection.protected`, используя их.</span><span class="sxs-lookup"><span data-stu-id="fe540-154">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="fe540-155">Только после этого станет возможным вызов `sheet.protection.protected` без исключения.</span><span class="sxs-lookup"><span data-stu-id="fe540-155">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="fe540-156">Процесс получения делится на три этапа:</span><span class="sxs-lookup"><span data-stu-id="fe540-156">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="fe540-157">Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.</span><span class="sxs-lookup"><span data-stu-id="fe540-157">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="fe540-158">Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.</span><span class="sxs-lookup"><span data-stu-id="fe540-158">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="fe540-159">Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.</span><span class="sxs-lookup"><span data-stu-id="fe540-159">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="fe540-160">Эти три действия должны выполняться каждый раз, когда коду нужно *прочесть* данные из документа Office.</span><span class="sxs-lookup"><span data-stu-id="fe540-160">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="fe540-p112">В функции `toggleProtection` замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="fe540-p112">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="fe540-163">У каждого объекта Excel есть метод `load`.</span><span class="sxs-lookup"><span data-stu-id="fe540-163">Every Excel object has a `load` method.</span></span> <span data-ttu-id="fe540-164">Вы указываете свойства объекта, которые нужно прочесть в параметре как строку имен, разделенных запятыми.</span><span class="sxs-lookup"><span data-stu-id="fe540-164">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="fe540-165">В этом случае нужно прочесть подсвойство свойства `protection`.</span><span class="sxs-lookup"><span data-stu-id="fe540-165">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="fe540-166">На подсвойство нужно ссылаться почти так же, как и в остальных частях кода. Отличие заключается в том, что вместо символа "." нужно указать косую черту ("/").</span><span class="sxs-lookup"><span data-stu-id="fe540-166">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>
   - <span data-ttu-id="fe540-167">Чтобы логика переключения, которая считывает `sheet.protection.protected`, не срабатывала до выполнения `sync` и присвоения `sheet.protection.protected` правильного значения, полученного из документа, она будет перемещена (на следующем этапе) в функцию `then`, которая не выполняется до завершения `sync`.</span><span class="sxs-lookup"><span data-stu-id="fe540-167">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="fe540-168">Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="fe540-168">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="fe540-169">Вы добавите новую последнюю строку `context.sync` позже.</span><span class="sxs-lookup"><span data-stu-id="fe540-169">You will add a new final `context.sync`, in a later step.</span></span>
3. <span data-ttu-id="fe540-170">Вырежьте структуру `if ... else` в функции `toggleProtection` и вставьте вместо `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="fe540-170">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>
4. <span data-ttu-id="fe540-p115">Замените `TODO4` приведенным ниже кодом. Примечание:</span><span class="sxs-lookup"><span data-stu-id="fe540-p115">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="fe540-173">Благодаря тому, что метод `sync` передается функции `then`, он не будет запускаться до добавления `sheet.protection.unprotect()` или `sheet.protection.protect()` в очередь.</span><span class="sxs-lookup"><span data-stu-id="fe540-173">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>
   - <span data-ttu-id="fe540-174">Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" после `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="fe540-174">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```javascript
    .then(context.sync);
    ```

   <span data-ttu-id="fe540-175">Когда все будет готово, функция должна выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="fe540-175">When you are done, the entire function should look like the following:</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="fe540-176">Настройка HTML-файла для загрузки скрипта</span><span class="sxs-lookup"><span data-stu-id="fe540-176">Configure the script-loading HTML file</span></span>

<span data-ttu-id="fe540-177">Откройте файл /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="fe540-177">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="fe540-178">Это HTML-файл без пользовательского интерфейса, вызываемый, когда пользователь нажимает кнопку **Toggle Worksheet Protection** (Переключение защиты листа).</span><span class="sxs-lookup"><span data-stu-id="fe540-178">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="fe540-179">Он предназначен для загрузки метода JavaScript, который должен выполняться при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="fe540-179">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="fe540-180">Вы не будете изменять этот файл.</span><span class="sxs-lookup"><span data-stu-id="fe540-180">You are not going to change this file.</span></span> <span data-ttu-id="fe540-181">Просто обратите внимание на то, что второй тег `<script>` загружает functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="fe540-181">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="fe540-182">Файл function-file.html и загружаемый им файл function-file.js выполняются в полностью отдельном процессе IE из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="fe540-182">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="fe540-183">Если файл function-file.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения.</span><span class="sxs-lookup"><span data-stu-id="fe540-183">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="fe540-184">Кроме того, файл function-file.js не содержит код JavaScript, который не поддерживается в IE.</span><span class="sxs-lookup"><span data-stu-id="fe540-184">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="fe540-185">По этим двум причинам такая надстройка не передает файл function-file.js вообще.</span><span class="sxs-lookup"><span data-stu-id="fe540-185">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="fe540-186">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="fe540-186">Test the add-in</span></span>

1. <span data-ttu-id="fe540-187">Закройте все приложения Office, в том числе Excel.</span><span class="sxs-lookup"><span data-stu-id="fe540-187">Close all Office applications, including Excel.</span></span> 
2. <span data-ttu-id="fe540-188">Очистите кэш Office, удалив содержимое папки кэша.</span><span class="sxs-lookup"><span data-stu-id="fe540-188">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="fe540-189">Это необходимо, чтобы можно было полностью удалить старую версию надстройки из ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="fe540-189">This is necessary to completely clear the old version of the add-in from the host.</span></span> 
    - <span data-ttu-id="fe540-190">Для Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="fe540-190">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>
    - <span data-ttu-id="fe540-191">Для Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="fe540-191">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>
3. <span data-ttu-id="fe540-192">Если по той или иной причине ваш сервер не работает, в окне Git Bash или системной командной строке с поддержкой Node.JS перейдите к папке **Start** проекта и выполните команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="fe540-192">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="fe540-193">Повторную сборку проекта выполнять не нужно, так как единственный файл JavaScript, который вы изменили, не относится к сборке bundle.js.</span><span class="sxs-lookup"><span data-stu-id="fe540-193">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>
4. <span data-ttu-id="fe540-194">Используя новую версию измененного файла манифеста, повторите процесс загрузки неопубликованного приложения с помощью одного из указанных далее методов.</span><span class="sxs-lookup"><span data-stu-id="fe540-194">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="fe540-195">*Нужно перезаписать предыдущий экземпляр файла манифеста.*</span><span class="sxs-lookup"><span data-stu-id="fe540-195">*You should overwrite the previous copy of the manifest file.*</span></span>
    - <span data-ttu-id="fe540-196">Windows: [загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="fe540-196">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="fe540-197">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="fe540-197">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="fe540-198">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="fe540-198">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
7. <span data-ttu-id="fe540-199">Откройте любой лист в Excel.</span><span class="sxs-lookup"><span data-stu-id="fe540-199">Open any worksheet in Excel.</span></span>
8. <span data-ttu-id="fe540-p121">На ленте **Главная** нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа). Обратите внимание на то, что большинство элементов управления на ленте отключены (серые), как показано на приведенном ниже снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="fe540-p121">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 
9. <span data-ttu-id="fe540-202">Выберите ячейку, как если бы вы хотели изменить ее содержимое.</span><span class="sxs-lookup"><span data-stu-id="fe540-202">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="fe540-203">Появится сообщение об ошибке и защите листа.</span><span class="sxs-lookup"><span data-stu-id="fe540-203">You get an error telling you that the worksheet is protected.</span></span>
10. <span data-ttu-id="fe540-204">Нажмите кнопку **Toggle Worksheet Protection** (Переключение защиты листа) еще раз, и элементы управления включатся, после чего вы сможете изменить значения ячеек.</span><span class="sxs-lookup"><span data-stu-id="fe540-204">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Руководство по Excel: лента с включенной защитой](../images/excel-tutorial-ribbon-with-protection-on.png)
