---
ms.date: 08/25/2020
title: Настройка надстройки Excel для совместного использования среды выполнения браузера
ms.prod: excel
description: Настройте надстройку Excel, чтобы предоставить общий доступ к среде выполнения браузера и запускать код ленты, области задач и пользовательских функций в одной и той же среде выполнения.
localization_priority: Priority
ms.openlocfilehash: 3f980ffc3ed78a4adf8c1b2cb565feb0f7c51c2f
ms.sourcegitcommit: 6ade8891ad947094d305fc146bb4deb703093ca6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/04/2020
ms.locfileid: "48906024"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="36446-103">Настройка надстройки Excel для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="36446-103">Configure your Excel add-in to use a shared JavaScript runtime</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="36446-104">При запуске Excel на компьютере с Windows или на Mac надстройка запустит код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="36446-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="36446-105">Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="36446-105">This creates limitations such as not being able to easily share global data, and not having access to all CORS functionality from a custom function.</span></span>

<span data-ttu-id="36446-106">Но вы можете настроить вашу надстройку Excel, предоставив общий доступ к коду в общей среде выполнения  JavaScript.</span><span class="sxs-lookup"><span data-stu-id="36446-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="36446-107">Это позволяет повысить слаженность работы всей вашей надстройки и обеспечить доступ к DOM и CORS из всех ее частей.</span><span class="sxs-lookup"><span data-stu-id="36446-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="36446-108">Кроме того, это позволяет запускать код при открытии документа и после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="36446-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="36446-109">Чтобы настроить надстройку для использования общей среды выполнения, следуйте инструкциям, приведенным в этой статье.</span><span class="sxs-lookup"><span data-stu-id="36446-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="36446-110">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="36446-110">Create the add-in project</span></span>

<span data-ttu-id="36446-111">Если вы начинаете новый проект, выполните указанные ниже действия, чтобы с помощью генератора Yeoman создать проект надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="36446-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="36446-112">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="36446-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="36446-113">Выберите тип проекта: **проект надстройки пользовательских функций Excel**</span><span class="sxs-lookup"><span data-stu-id="36446-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="36446-114">Выберите тип сценария: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="36446-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="36446-115">Как вы хотите назвать надстройку? **Моя надстройка Office**</span><span class="sxs-lookup"><span data-stu-id="36446-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

<span data-ttu-id="36446-117">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="36446-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="36446-118">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="36446-118">Configure the manifest</span></span>

<span data-ttu-id="36446-119">Выполните указанные ниже действия для нового или существующего проекта, чтобы настроить его для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="36446-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="36446-120">Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="36446-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="36446-121">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="36446-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="36446-122">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="36446-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="36446-123">Время существования должно быть **длительным** , чтобы пользовательские функции могли работать даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="36446-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="36446-124">Атрибут resid равен `ContosoAddin.Url` и ссылается на строку в разделе ресурсов далее.</span><span class="sxs-lookup"><span data-stu-id="36446-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="36446-125">Можно использовать любое значение resid, но оно должно соответствовать resid других элементов в элементах вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="36446-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="36446-126">В элементе `<Page>` замените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="36446-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="36446-127">Этот resid соответствует элементу resid `<Runtime>`.</span><span class="sxs-lookup"><span data-stu-id="36446-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="36446-128">Обратите внимание: если у вас нет пользовательских функций, то у вас не будет элемента **Page** , и этот шаг можно пропустить.</span><span class="sxs-lookup"><span data-stu-id="36446-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="36446-129">В разделе `<DesktopFormFactor>` замените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="36446-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="36446-130">Обратите внимание: если у вас нет команд действий, то у вас не будет элемента **FunctionFile** , и этот шаг можно пропустить.</span><span class="sxs-lookup"><span data-stu-id="36446-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="36446-131">В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="36446-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="36446-132">Обратите внимание: если у вас нет области задач, то у вас не будет действия **ShowTaskpane** , и этот шаг можно пропустить.</span><span class="sxs-lookup"><span data-stu-id="36446-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="36446-133">Добавьте новый **Url-идентификатор** для **ContosoAddin.Url** , указывающий на **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="36446-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/dist/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="36446-134">Убедитесь в том, что в файле taskpane.html есть тег `<script>`, ссылающийся на файл dist/functions.js.</span><span class="sxs-lookup"><span data-stu-id="36446-134">Make sure the taskpane.html has a `<script>` tag that references the dist/functions.js file.</span></span> <span data-ttu-id="36446-135">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="36446-135">The following is an example.</span></span>

   ```html
   <script type="text/javascript" src="/dist/functions.js" ></script>
   ```

   > [!NOTE]
   > <span data-ttu-id="36446-136">Если для вставки тегов сценариев надстройка использует Webpack и HtmlWebpackPlugin, как это делают надстройки, созданные генератором Yeoman (см. раздел [Создание проекта надстройки](#create-the-add-in-project) выше), то вам необходимо обеспечить включение модуля functions.js в массив `chunks`, как в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="36446-136">If the add-in uses Webpack and the HtmlWebpackPlugin to insert script tags, as add-ins created by the Yeoman generator do (see [Create the add-in project](#create-the-add-in-project) above), then you must ensure that the functions.js module is included in the `chunks` array as in the following example.</span></span>
   >
   > ```javascript
   > new HtmlWebpackPlugin({
   >     filename: "taskpane.html",
   >     template: "./src/taskpane/taskpane.html",
   >     chunks: ["polyfill", "taskpane", "functions"]
   > }),
   >```

9. <span data-ttu-id="36446-137">Сохраните изменения и перестройте проект.</span><span class="sxs-lookup"><span data-stu-id="36446-137">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="36446-138">Срок существования среды выполнения</span><span class="sxs-lookup"><span data-stu-id="36446-138">Runtime lifetime</span></span>

<span data-ttu-id="36446-139">Добавляя элемент `Runtime`, вы также задаете срок существования со значением `long` или `short`.</span><span class="sxs-lookup"><span data-stu-id="36446-139">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="36446-140">Установите значение `long`, чтобы воспользоваться такими функциями, как запуск надстройки при открытии документа, продолжение выполнения кода после закрытия области задач или использование CORS и DOM из пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="36446-140">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

>[!NOTE]
> <span data-ttu-id="36446-141">По умолчанию используется значение срока жизни `short`, но мы рекомендуем использовать `long` в надстройках Excel. Если вы настроите в этом примере для среды выполнения значение `short`, ваша надстройка Excel запустится при нажатии одной из кнопок на ленте, но может завершить работу после окончания функционирования обработчика ленты.</span><span class="sxs-lookup"><span data-stu-id="36446-141">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="36446-142">Точно так же, надстройка запустится при открытии области задач, но может завершить работу после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="36446-142">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> <span data-ttu-id="36446-143">Если в манифесте надстройки есть элемент `Runtimes` (требуемый для общей среды выполнения), она использует Internet Explorer 11 независимо от того, какая у вас версия Windows или Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="36446-143">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="36446-144">Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).</span><span class="sxs-lookup"><span data-stu-id="36446-144">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="multiple-task-panes"></a><span data-ttu-id="36446-145">Несколько областей задач</span><span class="sxs-lookup"><span data-stu-id="36446-145">Multiple task panes</span></span>

<span data-ttu-id="36446-146">Не планируйте использовать в своей надстройке несколько областей задач, если предполагается использование общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="36446-146">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="36446-147">Общая среда выполнения поддерживает только одну область задач.</span><span class="sxs-lookup"><span data-stu-id="36446-147">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="36446-148">Обратите внимание: любая область задач без `<TaskpaneID>` считается другой областью задач.</span><span class="sxs-lookup"><span data-stu-id="36446-148">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="36446-149">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="36446-149">Next steps</span></span>

- <span data-ttu-id="36446-150">Подробные сведения об использовании API JavaScript для Excel и пользовательских функций Excel в общей среде выполнения см. в статье [Вызов API Excel из пользовательской функции](call-excel-apis-from-custom-function.md).</span><span class="sxs-lookup"><span data-stu-id="36446-150">Read the [Call Excel APIs from a custom function](call-excel-apis-from-custom-function.md) article for details on using the Excel JavaScript APIs and custom Excel functions in a shared runtime.</span></span>
- <span data-ttu-id="36446-151">Изучите пример PnP [Управление интерфейсом ленты и области задач, а также запуск кода при открытии документа](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario), чтобы ознакомиться с масштабным примером работы общей среды выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="36446-151">Explore the patterns-and-practices sample [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) to see a larger example of the shared JavaScript runtime in action.</span></span>

## <a name="see-also"></a><span data-ttu-id="36446-152">См. также</span><span class="sxs-lookup"><span data-stu-id="36446-152">See also</span></span>

- [<span data-ttu-id="36446-153">Обзор: запуск кода надстройки в общей среде выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="36446-153">Overview: Run your add-in code in a shared JavaScript runtime</span></span>](custom-functions-shared-overview.md)
