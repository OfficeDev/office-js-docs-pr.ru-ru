---
ms.date: 04/08/2021
title: Настройка надстройки Office для использования общей среды выполнения JavaScript
ms.prod: non-product-specific
description: Настройте надстройку Office для использования общей среды выполнения JavaScript, чтобы применять дополнительные возможности ленты, области задач и пользовательских функций.
localization_priority: Priority
ms.openlocfilehash: d5f0a5b6d9053f23792012f1658d213a7972b970
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652196"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="9e893-103">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="9e893-103">Configure your Office Add-in to use a shared JavaScript runtime</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="9e893-104">Вы можете настроить надстройку Office, чтобы выполнять весь ее код в единой общей среде выполнения JavaScript (также называемой общей средой выполнения).</span><span class="sxs-lookup"><span data-stu-id="9e893-104">You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime).</span></span> <span data-ttu-id="9e893-105">Это позволяет повысить слаженность работы всей вашей надстройки и обеспечить доступ к DOM и CORS из всех ее частей.</span><span class="sxs-lookup"><span data-stu-id="9e893-105">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="9e893-106">Кроме того, это позволяет использовать дополнительные функции, например запуск кода при открытии документа, а также включение и отключение кнопок ленты.</span><span class="sxs-lookup"><span data-stu-id="9e893-106">It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons.</span></span> <span data-ttu-id="9e893-107">Чтобы настроить надстройку для использования общей среды выполнения JavaScript, следуйте инструкциям, приведенным в этой статье.</span><span class="sxs-lookup"><span data-stu-id="9e893-107">To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="9e893-108">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="9e893-108">Create the add-in project</span></span>

<span data-ttu-id="9e893-109">Если вы начинаете новый проект, выполните указанные ниже действия, чтобы с помощью [генератора Yeoman для настроек Office](https://github.com/OfficeDev/generator-office) создать проект надстройки Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9e893-109">If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.</span></span>

<span data-ttu-id="9e893-110">Выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="9e893-110">Do one of the following:</span></span>

- <span data-ttu-id="9e893-111">Чтобы создать надстройку Excel с пользовательскими функциями, выполните команду `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.</span><span class="sxs-lookup"><span data-stu-id="9e893-111">To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.</span></span>

    <span data-ttu-id="9e893-112">или</span><span class="sxs-lookup"><span data-stu-id="9e893-112">or</span></span>

- <span data-ttu-id="9e893-113">Чтобы создать надстройку PowerPoint, выполните команду `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.</span><span class="sxs-lookup"><span data-stu-id="9e893-113">To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.</span></span>

<span data-ttu-id="9e893-114">Генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="9e893-114">The generator will create the project and install supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="9e893-115">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="9e893-115">Configure the manifest</span></span>

<span data-ttu-id="9e893-116">Выполните указанные ниже действия для нового или существующего проекта, чтобы настроить его для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="9e893-116">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span> <span data-ttu-id="9e893-117">Эти действия подразумевают, что вы создали проект с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="9e893-117">These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="9e893-118">Запустите код Visual Studio и откройте созданный вами проект надстройки Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9e893-118">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="9e893-119">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="9e893-119">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="9e893-120">Если вы создали надстройку для Excel, обновите раздел требований, чтобы использовать [общую среду выполнения](../reference/requirement-sets/shared-runtime-requirement-sets.md), а не среду выполнения пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="9e893-120">If you generated an Excel add-in, update the requirements section to use the [shared runtime](../reference/requirement-sets/shared-runtime-requirement-sets.md) instead of the custom function runtime.</span></span> <span data-ttu-id="9e893-121">XML-код должен выглядеть следующим образом.</span><span class="sxs-lookup"><span data-stu-id="9e893-121">The XML should appear as follows.</span></span>

    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
    </Requirements>
    ```

1. <span data-ttu-id="9e893-122">Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>` внутри тега `<Host ...>`.</span><span class="sxs-lookup"><span data-stu-id="9e893-122">Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag.</span></span> <span data-ttu-id="9e893-123">Время существования должно иметь значение **long**, чтобы код надстройки мог выполняться даже после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="9e893-123">The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed.</span></span> <span data-ttu-id="9e893-124">Значение `resid` — **Taskpane.Url**, указывающее расположение файла **taskpane.html** в разделе ` <bt:Urls>` в нижней части **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="9e893-124">The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.</span></span>

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       ...
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. <span data-ttu-id="9e893-125">Если вы создали надстройку Excel с пользовательскими функциями, найдите элемент `<Page>`.</span><span class="sxs-lookup"><span data-stu-id="9e893-125">If you generated an Excel add-in with custom functions, find the `<Page>` element.</span></span> <span data-ttu-id="9e893-126">Затем измените расположение источника с **Functions.Page.Url** на **Taskpane.Url**.</span><span class="sxs-lookup"><span data-stu-id="9e893-126">Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. <span data-ttu-id="9e893-127">Найдите тег `<FunctionFile ...>` и измените `resid` с **Commands.Url** на **Taskpane.Url**.</span><span class="sxs-lookup"><span data-stu-id="9e893-127">Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.</span></span> <span data-ttu-id="9e893-128">Обратите внимание: если у вас нет команд действий, у вас не будет записи **FunctionFile**, и этот шаг можно пропустить.</span><span class="sxs-lookup"><span data-stu-id="9e893-128">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. <span data-ttu-id="9e893-129">Сохраните файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="9e893-129">Save the **manifest.xml** file.</span></span>

## <a name="configure-the-webpackconfigjs-file"></a><span data-ttu-id="9e893-130">Настройка файла webpack.config.js.</span><span class="sxs-lookup"><span data-stu-id="9e893-130">Configure the webpack.config.js file</span></span>

<span data-ttu-id="9e893-131">Файл **webpack.config.js** создает несколько загрузчиков среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="9e893-131">The **webpack.config.js** will build multiple runtime loaders.</span></span> <span data-ttu-id="9e893-132">Вам требуется изменить его, чтобы загружать только общую среду выполнения JavaScript с помощью файла **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="9e893-132">You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.</span></span>

1. <span data-ttu-id="9e893-133">Запустите код Visual Studio и откройте созданный вами проект надстройки Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9e893-133">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
1. <span data-ttu-id="9e893-134">Откройте файл **webpack.config.js**.</span><span class="sxs-lookup"><span data-stu-id="9e893-134">Open the **webpack.config.js** file.</span></span>
1. <span data-ttu-id="9e893-135">Если файл **webpack.config.js** содержит следующий код подключаемого модуля **functions.html**, удалите его.</span><span class="sxs-lookup"><span data-stu-id="9e893-135">If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. <span data-ttu-id="9e893-136">Если файл **webpack.config.js** содержит следующий код подключаемого модуля **commands.html**, удалите его.</span><span class="sxs-lookup"><span data-stu-id="9e893-136">If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. <span data-ttu-id="9e893-137">Если в проекте используются блоки **functions** или **commands**, добавьте их в список блоков, как показано ниже (следующий код предназначен для проекта, применяющего оба блока).</span><span class="sxs-lookup"><span data-stu-id="9e893-137">If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).</span></span>

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. <span data-ttu-id="9e893-138">Сохраните изменения и выполните повторную сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="9e893-138">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

> [!NOTE]
> <span data-ttu-id="9e893-139">Если в проекте есть файлы **functions.html** или **commands.html**, их можно удалить.</span><span class="sxs-lookup"><span data-stu-id="9e893-139">If your project has a **functions.html** file or **commands.html** file, they can be removed.</span></span> <span data-ttu-id="9e893-140">**Taskpane.html** загружает код **functions.js** и **commands.js** в общую среду выполнения JavaScript с помощью созданных вами обновлений webpack.</span><span class="sxs-lookup"><span data-stu-id="9e893-140">The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.</span></span>

## <a name="test-your-office-add-in-changes"></a><span data-ttu-id="9e893-141">Тестирование изменений надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9e893-141">Test your Office Add-in changes</span></span>

<span data-ttu-id="9e893-142">Вы можете убедиться, что вы используете общую среду выполнения JavaScript надлежащим образом, воспользовавшись следующими инструкциями.</span><span class="sxs-lookup"><span data-stu-id="9e893-142">You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.</span></span>

1. <span data-ttu-id="9e893-143">Откройте файл **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="9e893-143">Open the **manifest.xml** file.</span></span>
1. <span data-ttu-id="9e893-144">Найдите раздел `<Control xsi:type="Button" id="TaskpaneButton">` и измените следующий XML-код `<Action ...>`.</span><span class="sxs-lookup"><span data-stu-id="9e893-144">Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.</span></span>

    <span data-ttu-id="9e893-145">с:</span><span class="sxs-lookup"><span data-stu-id="9e893-145">from:</span></span>

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    <span data-ttu-id="9e893-146">на:</span><span class="sxs-lookup"><span data-stu-id="9e893-146">to:</span></span>

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. <span data-ttu-id="9e893-147">Откройте файл **./src/commands/commands.js**.</span><span class="sxs-lookup"><span data-stu-id="9e893-147">Open the **./src/commands/commands.js** file.</span></span>
1. <span data-ttu-id="9e893-148">Замените имеющуюся функцию **action** указанным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9e893-148">Replace the **action** function with the code below.</span></span> <span data-ttu-id="9e893-149">При этом функция будет обновлена для открытия и изменения кнопки области задач, чтобы увеличить счетчик.</span><span class="sxs-lookup"><span data-stu-id="9e893-149">This will update the function to open and modify the task pane button to increment a counter.</span></span> <span data-ttu-id="9e893-150">Открытие модели DOM области задач и доступ к ней из команды поддерживается только в общей среде выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9e893-150">Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.</span></span>

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. <span data-ttu-id="9e893-151">Сохраните изменения и запустите проект.</span><span class="sxs-lookup"><span data-stu-id="9e893-151">Save your changes and run the project.</span></span>

   ```command line
   npm start
   ```

<span data-ttu-id="9e893-152">Каждый раз при нажатии кнопки надстройки текст кнопки **run** (выполнить) будет изменяться на **go** (перейти) с увеличением счетчика после этого.</span><span class="sxs-lookup"><span data-stu-id="9e893-152">Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.</span></span>

## <a name="runtime-lifetime"></a><span data-ttu-id="9e893-153">Срок существования среды выполнения</span><span class="sxs-lookup"><span data-stu-id="9e893-153">Runtime lifetime</span></span>

<span data-ttu-id="9e893-154">Добавляя элемент `Runtime`, вы также задаете срок существования со значением `long` или `short`.</span><span class="sxs-lookup"><span data-stu-id="9e893-154">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="9e893-155">Установите значение `long`, чтобы воспользоваться такими функциями, как запуск надстройки при открытии документа, продолжение выполнения кода после закрытия области задач или использование CORS и DOM из пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="9e893-155">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

> [!NOTE]
> <span data-ttu-id="9e893-156">По умолчанию используется значение срока жизни `short`, но мы рекомендуем использовать `long` в надстройках Excel. Если вы настроите в этом примере для среды выполнения значение `short`, ваша надстройка Excel запустится при нажатии одной из кнопок на ленте, но может завершить работу после окончания функционирования обработчика ленты.</span><span class="sxs-lookup"><span data-stu-id="9e893-156">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="9e893-157">Аналогичным образом надстройка запустится при открытии области задач, но может завершить работу после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="9e893-157">Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> <span data-ttu-id="9e893-158">Если в манифесте надстройки есть элемент `Runtimes` (требуемый для общей среды выполнения), она использует Internet Explorer 11 независимо от того, какая у вас версия Windows или Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="9e893-158">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="9e893-159">Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).</span><span class="sxs-lookup"><span data-stu-id="9e893-159">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="about-the-shared-javascript-runtime"></a><span data-ttu-id="9e893-160">Сведения об общей среде выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="9e893-160">About the shared JavaScript runtime</span></span>

<span data-ttu-id="9e893-161">На компьютере с Windows или Mac надстройка запускает код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9e893-161">On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="9e893-162">Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="9e893-162">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="9e893-163">Однако вы можете настроить надстройку Office так, чтобы обеспечить общий доступ к коду в одной среде выполнения JavaScript (то есть в общей среде выполнения).</span><span class="sxs-lookup"><span data-stu-id="9e893-163">However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="9e893-164">За счет этого повышается скоординированность работы надстройки и упрощается доступ к модели DOM и CORS области задач из всех компонентов надстройки.</span><span class="sxs-lookup"><span data-stu-id="9e893-164">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="9e893-165">При настройке общей среды выполнения становятся возможными следующие сценарии.</span><span class="sxs-lookup"><span data-stu-id="9e893-165">Configuring a shared runtime enables the following scenarios.</span></span>

- <span data-ttu-id="9e893-166">Надстройка Office может использовать дополнительные функции пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="9e893-166">Your Office Add-in can use additional UI features:</span></span>
  - [<span data-ttu-id="9e893-167">Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9e893-167">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
  - [<span data-ttu-id="9e893-168">Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9e893-168">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
  - [<span data-ttu-id="9e893-169">Включение и отключение команд надстроек</span><span class="sxs-lookup"><span data-stu-id="9e893-169">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
  - [<span data-ttu-id="9e893-170">Запуск кода в надстройке Office при открытии документа</span><span class="sxs-lookup"><span data-stu-id="9e893-170">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
  - [<span data-ttu-id="9e893-171">Отображение и скрытие области задач надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9e893-171">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- <span data-ttu-id="9e893-172">Для надстроек Excel:</span><span class="sxs-lookup"><span data-stu-id="9e893-172">For Excel add-ins:</span></span>
  - <span data-ttu-id="9e893-173">Пользовательские функции полностью поддерживают CORS.</span><span class="sxs-lookup"><span data-stu-id="9e893-173">Custom functions will have full CORS support.</span></span>
  - <span data-ttu-id="9e893-174">Пользовательские функции могут вызывать API Office.js для чтения данных из электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="9e893-174">Custom functions can call Office.js APIs to read spreadsheet document data.</span></span>

<span data-ttu-id="9e893-175">Для Office в Windows общая среда выполнения требует наличия экземпляра браузера Microsoft Internet Explorer 11, как описано в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). Кроме того, все кнопки, отображаемые вашей надстройкой на ленте, будут работать в этой же общей среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="9e893-175">For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="9e893-176">На следующем рисунке показано, как пользовательские функции, пользовательский интерфейс ленты и код области задач будут запускаться в одной среде выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9e893-176">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Схема пользовательской функции, области задач и кнопок ленты, работающих в общей среде выполнения браузера IE в Excel](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a><span data-ttu-id="9e893-178">Отладка</span><span class="sxs-lookup"><span data-stu-id="9e893-178">Debugging</span></span>

<span data-ttu-id="9e893-179">В настоящее время при использовании общей среды выполнения невозможно использовать Visual Studio Code для отладки пользовательских функций в Excel под управлением Windows.</span><span class="sxs-lookup"><span data-stu-id="9e893-179">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="9e893-180">Вместо этого потребуется использовать средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="9e893-180">You'll need to use developer tools instead.</span></span> <span data-ttu-id="9e893-181">Дополнительные сведения см. в статье [Отладка надстроек с помощью средств разработчика в Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="9e893-181">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

### <a name="multiple-task-panes"></a><span data-ttu-id="9e893-182">Несколько областей задач</span><span class="sxs-lookup"><span data-stu-id="9e893-182">Multiple task panes</span></span>

<span data-ttu-id="9e893-183">Не планируйте использовать в своей надстройке несколько областей задач, если предполагается использование общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="9e893-183">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="9e893-184">Общая среда выполнения поддерживает только одну область задач.</span><span class="sxs-lookup"><span data-stu-id="9e893-184">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="9e893-185">Обратите внимание: любая область задач без `<TaskpaneID>` считается другой областью задач.</span><span class="sxs-lookup"><span data-stu-id="9e893-185">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="9e893-186">Напишите нам свой отзыв</span><span class="sxs-lookup"><span data-stu-id="9e893-186">Give us feedback</span></span>

<span data-ttu-id="9e893-187">Мы будем рады услышать ваши отзывы об этой функции.</span><span class="sxs-lookup"><span data-stu-id="9e893-187">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="9e893-188">Если вы обнаружите какие-либо ошибки или проблемы, если у вас есть запросы относительно этой функции, сообщите нам, создав проблему GitHub в [репозитории office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="9e893-188">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="9e893-189">См. также</span><span class="sxs-lookup"><span data-stu-id="9e893-189">See also</span></span>

- [<span data-ttu-id="9e893-190">Вызов API Excel из пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="9e893-190">Call Excel APIs from a custom function</span></span>](../excel/call-excel-apis-from-custom-function.md)
- [<span data-ttu-id="9e893-191">Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9e893-191">Add custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
- [<span data-ttu-id="9e893-192">Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9e893-192">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
- [<span data-ttu-id="9e893-193">Включение и отключение команд надстроек</span><span class="sxs-lookup"><span data-stu-id="9e893-193">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
- [<span data-ttu-id="9e893-194">Запуск кода в надстройке Office при открытии документа</span><span class="sxs-lookup"><span data-stu-id="9e893-194">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
- [<span data-ttu-id="9e893-195">Отображение и скрытие области задач надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9e893-195">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- [<span data-ttu-id="9e893-196">Учебное руководство. Обмен данными и событиями между пользовательскими функциями Excel и областью задач</span><span class="sxs-lookup"><span data-stu-id="9e893-196">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
