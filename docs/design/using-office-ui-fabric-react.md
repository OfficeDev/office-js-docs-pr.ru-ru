---
title: Использование Office UI Fabric React в надстройках Office
description: Использование Office UI Fabric React в надстройках Office
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: f8f61d1b094fa71b8a400a6a6d9ea3029c53b051
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237730"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="51aed-103">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="51aed-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="51aed-p101">Office UI Fabric — это интерфейсная структура JavaScript для создания пользовательского интерфейса для Office. Если вы создаете надстройки с помощью React, рассмотрите возможность использования Fabric React для создания пользовательского интерфейса. Fabric предоставляет несколько компонентов UX на основе React, например кнопки или контрольные элементы, которые можно использовать в надстройке.</span><span class="sxs-lookup"><span data-stu-id="51aed-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="51aed-107">В этой статье объясняется, как создать надстройку с помощью React и использованием компонентов Fabric React.</span><span class="sxs-lookup"><span data-stu-id="51aed-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span>

> [!NOTE]
> <span data-ttu-id="51aed-108">В Fabric React используется[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors), поэтому после выполнения вами действий, указанных в этой статье, ваша надстройка будет включать и доступ к Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="51aed-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="51aed-109">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="51aed-109">Create an add-in project</span></span>

<span data-ttu-id="51aed-110">Чтобы создать надстройку с использованием React, рекомендуется воспользоваться генератором Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="51aed-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="51aed-111">Установка необходимых компонентов</span><span class="sxs-lookup"><span data-stu-id="51aed-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="51aed-112">Создание проекта</span><span class="sxs-lookup"><span data-stu-id="51aed-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="51aed-113">**Выберите тип проекта:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="51aed-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="51aed-114">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="51aed-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="51aed-115">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="51aed-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="51aed-116">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="51aed-116">**Which Office client application would you like to support?**</span></span> `Word`

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки](../images/yo-office-word-react.png)

<span data-ttu-id="51aed-118">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="51aed-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="51aed-119">Проверка</span><span class="sxs-lookup"><span data-stu-id="51aed-119">Try it out</span></span>

1. <span data-ttu-id="51aed-120">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="51aed-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="51aed-121">Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="51aed-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="51aed-122">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="51aed-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="51aed-123">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="51aed-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="51aed-124">Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.</span><span class="sxs-lookup"><span data-stu-id="51aed-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="51aed-125">Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="51aed-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="51aed-126">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="51aed-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="51aed-127">Чтобы проверить надстройку в Word, выполните приведенную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="51aed-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="51aed-128">При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Word с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="51aed-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="51aed-129">Чтобы проверить надстройку в Word в браузере, выполните приведенную ниже команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="51aed-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="51aed-130">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="51aed-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="51aed-131">Чтобы использовать надстройку, откройте новый документ в Word в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="51aed-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="51aed-132">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="51aed-132">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="51aed-133">Обратите внимание на текст по умолчанию и кнопку **Запустить** в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="51aed-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="51aed-134">Следуя этой инструкции до конца, вы переопределите эти текст и кнопку, создав компонент React с использованием компонентов дизайна Fabric React.</span><span class="sxs-lookup"><span data-stu-id="51aed-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![Screenshot showing the Word application with the Show Taskpane ribbon button highlighted and the Run button and immediately preceding text highlighted in the task pane](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="51aed-136">Создание компонента React c использованием Fabric React</span><span class="sxs-lookup"><span data-stu-id="51aed-136">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="51aed-137">На этом этапе вы уже создали самую простую надстройку в области задач c использованием React.</span><span class="sxs-lookup"><span data-stu-id="51aed-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="51aed-138">Теперь выполните приведенные ниже действия, чтобы создать новый компонент React (`ButtonPrimaryExample`) в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="51aed-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="51aed-139">В этом компоненте будут использованы компоненты `Label` и `PrimaryButton` из Fabric React.</span><span class="sxs-lookup"><span data-stu-id="51aed-139">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="51aed-140">Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="51aed-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="51aed-141">Создайте в этой папке новый файл под названием **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="51aed-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="51aed-142">Введите в файл **Button.tsx** приведенный ниже код, чтобы определить компонент `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="51aed-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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
      body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
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

<span data-ttu-id="51aed-143">Этот код выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="51aed-143">This code does the following:</span></span>

- <span data-ttu-id="51aed-144">Ссылается на библиотеку React с помощью `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="51aed-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="51aed-145">Ссылается на компоненты Fabric (`PrimaryButton`, `IButtonProps`, `Label`), которые используются для создания `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="51aed-145">References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="51aed-146">Объявляет новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="51aed-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="51aed-147">Объявляет функцию `insertText` для обработки события кнопки `onClick`.</span><span class="sxs-lookup"><span data-stu-id="51aed-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="51aed-148">Определяет пользовательский интерфейс компонента React в функции `render`.</span><span class="sxs-lookup"><span data-stu-id="51aed-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="51aed-149">В HTML-разметке используются компоненты `Label` и `PrimaryButton` из Fabric React и указывается, что при подключения события `onClick` будет запускаться функция `insertText`.</span><span class="sxs-lookup"><span data-stu-id="51aed-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="51aed-150">Добавление компонента React в надстройку</span><span class="sxs-lookup"><span data-stu-id="51aed-150">Add the React component to your add-in</span></span>

<span data-ttu-id="51aed-151">Добавьте компонент `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\App.tsx** и выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="51aed-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="51aed-152">Добавьте приведенный ниже оператор импорта для ссылки на `ButtonPrimaryExample` из **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="51aed-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="51aed-153">Удалите два приведенные ниже оператора импорта.</span><span class="sxs-lookup"><span data-stu-id="51aed-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="51aed-154">Замените функцию по умолчанию `render()` на приведенный ниже код, в котором используется `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="51aed-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="51aed-155">Сохраните изменения, внесенные в **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="51aed-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="51aed-156">Результат</span><span class="sxs-lookup"><span data-stu-id="51aed-156">See the result</span></span>

<span data-ttu-id="51aed-157">После сохранения изменений в **App.tsx** область задач надстройки в Word обновляется автоматически. </span><span class="sxs-lookup"><span data-stu-id="51aed-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="51aed-158">Текст по умолчанию и кнопка в нижней части области задач теперь отображают пользовательский интерфейс, определяемый компонентом `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="51aed-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="51aed-159">Нажмите кнопку **Вставить текст...** для вставки текста в документ.</span><span class="sxs-lookup"><span data-stu-id="51aed-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![Screenshot showing the Word application with the "Insert text..." кнопка и сразу после выделенного текста](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="51aed-161">Поздравляем! Вы успешно создали надстройку области задач с помощью React и Office UI Fabric React!</span><span class="sxs-lookup"><span data-stu-id="51aed-161">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span>

## <a name="see-also"></a><span data-ttu-id="51aed-162">См. также</span><span class="sxs-lookup"><span data-stu-id="51aed-162">See also</span></span>

- [<span data-ttu-id="51aed-163">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="51aed-163">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="51aed-164">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="51aed-164">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="51aed-165">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="51aed-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="51aed-166">Начало работы с примером кода Fabric React</span><span class="sxs-lookup"><span data-stu-id="51aed-166">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
