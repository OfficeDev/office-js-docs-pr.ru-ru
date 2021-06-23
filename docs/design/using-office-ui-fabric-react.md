---
title: Fluent UI React в надстройках Office
description: Узнайте, как использовать Fluent интерфейс React в Office надстройки.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: a71c1a0de64d99a9e52c4ca2a7a948b9c33eb9ed
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076310"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a><span data-ttu-id="03f43-103">Использование Fluent интерфейса React в Office надстройки</span><span class="sxs-lookup"><span data-stu-id="03f43-103">Use Fluent UI React in Office Add-ins</span></span>

<span data-ttu-id="03f43-104">Fluent Интерфейс React является официальной интерфейсной платформой JavaScript с открытым исходным кодом, предназначенной для создания интерфейсных интерфейсов, которые легко вписываются в широкий спектр продуктов Майкрософт, включая Office.</span><span class="sxs-lookup"><span data-stu-id="03f43-104">Fluent UI React is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Office.</span></span> <span data-ttu-id="03f43-105">Он обеспечивает надежные, современные, доступные компоненты на основе React, которые легко настраиваются с помощью CSS-in-JS.</span><span class="sxs-lookup"><span data-stu-id="03f43-105">It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.</span></span>

> [!NOTE]
> <span data-ttu-id="03f43-106">В этой статье описывается использование Fluent пользовательского React в контексте Office надстройки. Но он также используется в широком диапазоне Microsoft 365 приложений и расширений.</span><span class="sxs-lookup"><span data-stu-id="03f43-106">This article describes the use of Fluent UI React in the context of Office Add-ins. But it is also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="03f43-107">Дополнительные сведения см. [в Fluent веб React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) пользовательского интерфейса и репо с открытым исходным кодом [Fluent пользовательского интерфейса.](https://github.com/microsoft/fluentui)</span><span class="sxs-lookup"><span data-stu-id="03f43-107">For more information, see [Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo [Fluent UI Web](https://github.com/microsoft/fluentui).</span></span>

<span data-ttu-id="03f43-108">В этой статье описывается, как создать надстройку, созданную с React и использующую Fluent компоненты React пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="03f43-108">This article describes how to create an add-in that's built with React and uses Fluent UI React components.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="03f43-109">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="03f43-109">Create an add-in project</span></span>

<span data-ttu-id="03f43-110">Чтобы создать надстройку с использованием React, рекомендуется воспользоваться генератором Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="03f43-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="03f43-111">Установка необходимых компонентов</span><span class="sxs-lookup"><span data-stu-id="03f43-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="03f43-112">Создание проекта</span><span class="sxs-lookup"><span data-stu-id="03f43-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="03f43-113">**Выберите тип проекта:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="03f43-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="03f43-114">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="03f43-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="03f43-115">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="03f43-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="03f43-116">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="03f43-116">**Which Office client application would you like to support?**</span></span> `Word`

![Снимок экрана, показывающий подсказки и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-word-react.png)

<span data-ttu-id="03f43-118">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="03f43-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="03f43-119">Проверка</span><span class="sxs-lookup"><span data-stu-id="03f43-119">Try it out</span></span>

1. <span data-ttu-id="03f43-120">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="03f43-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="03f43-121">Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="03f43-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="03f43-122">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="03f43-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="03f43-123">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="03f43-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="03f43-124">Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.</span><span class="sxs-lookup"><span data-stu-id="03f43-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="03f43-125">Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="03f43-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="03f43-126">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="03f43-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="03f43-127">Чтобы проверить надстройку в Word, выполните приведенную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="03f43-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="03f43-128">При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Word с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="03f43-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="03f43-129">Чтобы проверить надстройку в Word в браузере, выполните приведенную ниже команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="03f43-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="03f43-130">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="03f43-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="03f43-131">Чтобы использовать надстройку, откройте новый документ в Word в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="03f43-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="03f43-132">Чтобы открыть области задач надстройки, на вкладке **Главная** выберите кнопку **Показать задачу.**</span><span class="sxs-lookup"><span data-stu-id="03f43-132">To open the add-in task pane, on the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="03f43-133">Обратите внимание на текст по умолчанию и кнопку **Запустить** в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="03f43-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="03f43-134">В остальной части этого поголовия вы переопределяете этот текст и кнопку, создав компонент React, использующий компоненты UX из Fluent пользовательского React.</span><span class="sxs-lookup"><span data-stu-id="03f43-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fluent UI React.</span></span>

    ![Снимок экрана, показывающий приложение Word с выделенной кнопкой ленты Show Taskpane и кнопкой Run и непосредственно предшествующим текстом, выделенным в области задач.](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a><span data-ttu-id="03f43-136">Создайте компонент React, использующий Fluent пользовательского React</span><span class="sxs-lookup"><span data-stu-id="03f43-136">Create a React component that uses Fluent UI React</span></span>

<span data-ttu-id="03f43-137">На этом этапе вы уже создали самую простую надстройку в области задач c использованием React.</span><span class="sxs-lookup"><span data-stu-id="03f43-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="03f43-138">Теперь выполните приведенные ниже действия, чтобы создать новый компонент React (`ButtonPrimaryExample`) в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="03f43-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="03f43-139">Компонент использует `Label` компоненты из `PrimaryButton` Fluent пользовательского React.</span><span class="sxs-lookup"><span data-stu-id="03f43-139">The component uses the `Label` and `PrimaryButton` components from Fluent UI React.</span></span>

1. <span data-ttu-id="03f43-140">Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="03f43-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="03f43-141">Создайте в этой папке новый файл под названием **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="03f43-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="03f43-142">Введите в файл **Button.tsx** приведенный ниже код, чтобы определить компонент `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="03f43-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Fluent UI React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

<span data-ttu-id="03f43-143">Этот код выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="03f43-143">This code does the following:</span></span>

- <span data-ttu-id="03f43-144">Ссылается на библиотеку React с помощью `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="03f43-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="03f43-145">Ссылки Fluent пользовательского интерфейса React (, , ), которые `PrimaryButton` используются для создания `IButtonProps` `Label` `ButtonPrimaryExample` .</span><span class="sxs-lookup"><span data-stu-id="03f43-145">References the Fluent UI React components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="03f43-146">Объявляет новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="03f43-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="03f43-147">Объявляет функцию `insertText` для обработки события кнопки `onClick`.</span><span class="sxs-lookup"><span data-stu-id="03f43-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="03f43-148">Определяет пользовательский интерфейс компонента React в функции `render`.</span><span class="sxs-lookup"><span data-stu-id="03f43-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="03f43-149">HtmL-разметка использует компоненты Fluent пользовательского интерфейса React и указывает, что при запуске события функция `Label` `PrimaryButton` будет `onClick` `insertText` работать.</span><span class="sxs-lookup"><span data-stu-id="03f43-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fluent UI React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="03f43-150">Добавление компонента React в надстройку</span><span class="sxs-lookup"><span data-stu-id="03f43-150">Add the React component to your add-in</span></span>

<span data-ttu-id="03f43-151">Добавьте компонент `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\App.tsx** и выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="03f43-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="03f43-152">Добавьте приведенный ниже оператор импорта для ссылки на `ButtonPrimaryExample` из **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="03f43-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="03f43-153">Удалите два приведенные ниже оператора импорта.</span><span class="sxs-lookup"><span data-stu-id="03f43-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="03f43-154">Замените функцию по умолчанию `render()` на приведенный ниже код, в котором используется `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="03f43-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

    ```typescript
    render() {
      return (
        <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
          <ButtonPrimaryExample />
        </HeroList>
        </div>
      );
    }
    ```

4. <span data-ttu-id="03f43-155">Сохраните изменения, внесенные в **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="03f43-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="03f43-156">Результат</span><span class="sxs-lookup"><span data-stu-id="03f43-156">See the result</span></span>

<span data-ttu-id="03f43-157">После сохранения изменений в **App.tsx** область задач надстройки в Word обновляется автоматически. </span><span class="sxs-lookup"><span data-stu-id="03f43-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="03f43-158">Текст по умолчанию и кнопка в нижней части области задач теперь отображают пользовательский интерфейс, определяемый компонентом `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="03f43-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="03f43-159">Нажмите кнопку **Вставить текст...** для вставки текста в документ.</span><span class="sxs-lookup"><span data-stu-id="03f43-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Снимок экрана, показывающий приложение Word с текстом "Вставить текст...". кнопку и сразу перед выделенным текстом.](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="03f43-161">Поздравляем, вы успешно создали надстройку области задач с помощью React и Fluent пользовательского интерфейса React!</span><span class="sxs-lookup"><span data-stu-id="03f43-161">Congratulations, you've successfully created a task pane add-in using React and Fluent UI React!</span></span>

## <a name="see-also"></a><span data-ttu-id="03f43-162">См. также</span><span class="sxs-lookup"><span data-stu-id="03f43-162">See also</span></span>

- [<span data-ttu-id="03f43-163">Word Add-in GettingStartedFabricReact</span><span class="sxs-lookup"><span data-stu-id="03f43-163">Word Add-in GettingStartedFabricReact</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="03f43-164">Fabric Core в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="03f43-164">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="03f43-165">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="03f43-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
