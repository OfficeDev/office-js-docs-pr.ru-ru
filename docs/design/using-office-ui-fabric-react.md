---
title: Использование Office UI Fabric React в надстройках Office
description: Использование Office UI Fabric React в надстройках Office
ms.date: 01/16/2020
localization_priority: Normal
ms.openlocfilehash: 3891b3468b13823712afe93d0d1bb4d6d74faacb
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950462"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="b301a-103">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="b301a-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="b301a-p101">Office UI Fabric — это интерфейсная платформа JavaScript для построения взаимодействия с пользователем в Office и Office 365. Если вы разрабатываете надстройку с использованием React, пользовательский интерфейс рекомендуется создать с помощью Fabric React. В Fabric предоставлены некоторые компоненты дизайна на основе React, например кнопки и флажки, которые можно использовать в надстройке.</span><span class="sxs-lookup"><span data-stu-id="b301a-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="b301a-107">В этой статье объясняется, как создать надстройку с помощью React и использованием компонентов Fabric React.</span><span class="sxs-lookup"><span data-stu-id="b301a-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span> 

> [!NOTE]
> <span data-ttu-id="b301a-108">В Fabric React используется[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors), поэтому после выполнения вами действий, указанных в этой статье, ваша надстройка будет включать и доступ к Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="b301a-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="b301a-109">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="b301a-109">Create an add-in project</span></span>

<span data-ttu-id="b301a-110">Чтобы создать надстройку с использованием React, рекомендуется воспользоваться генератором Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="b301a-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="b301a-111">Установка необходимых компонентов</span><span class="sxs-lookup"><span data-stu-id="b301a-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="b301a-112">Создание проекта</span><span class="sxs-lookup"><span data-stu-id="b301a-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="b301a-113">**Выберите тип проекта:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="b301a-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="b301a-114">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="b301a-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="b301a-115">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="b301a-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="b301a-116">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="b301a-116">**Which Office client application would you like to support?**</span></span> `Word`

![Генератор Yeoman](../images/yo-office-word-react.png)

<span data-ttu-id="b301a-118">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="b301a-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="b301a-119">Проверка</span><span class="sxs-lookup"><span data-stu-id="b301a-119">Try it out</span></span>

1. <span data-ttu-id="b301a-120">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="b301a-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="b301a-121">Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="b301a-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b301a-122">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="b301a-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="b301a-123">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b301a-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="b301a-124">Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="b301a-124">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="b301a-125">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="b301a-125">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="b301a-126">Чтобы проверить надстройку в Word, выполните приведенную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="b301a-126">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="b301a-127">При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Word с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="b301a-127">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="b301a-128">Чтобы проверить надстройку в Word в браузере, выполните приведенную ниже команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="b301a-128">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="b301a-129">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="b301a-129">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="b301a-130">Чтобы использовать надстройку, откройте новый документ в Word в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="b301a-130">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="b301a-131">В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="b301a-131">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="b301a-132">Обратите внимание на текст по умолчанию и кнопку **Запустить** в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="b301a-132">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="b301a-133">Следуя этой инструкции до конца, вы переопределите эти текст и кнопку, создав компонент React с использованием компонентов дизайна Fabric React.</span><span class="sxs-lookup"><span data-stu-id="b301a-133">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![Снимок экрана c приложением Word с выделенными кнопками "Показать область задач", "Запустить" и предшествующим текстом в области задач](../images/word-task-pane-yo-default.png)


## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="b301a-135">Создание компонента React c использованием Fabric React</span><span class="sxs-lookup"><span data-stu-id="b301a-135">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="b301a-136">На этом этапе вы уже создали самую простую надстройку в области задач c использованием React.</span><span class="sxs-lookup"><span data-stu-id="b301a-136">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="b301a-137">Теперь выполните приведенные ниже действия, чтобы создать новый компонент React (`ButtonPrimaryExample`) в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="b301a-137">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="b301a-138">В этом компоненте будут использованы компоненты `Label` и `PrimaryButton` из Fabric React.</span><span class="sxs-lookup"><span data-stu-id="b301a-138">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="b301a-139">Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\taskpane\components**.</span><span class="sxs-lookup"><span data-stu-id="b301a-139">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="b301a-140">Создайте в этой папке новый файл под названием**Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="b301a-140">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="b301a-141">Введите в файл **Button.tsx** приведенный ниже код, чтобы определить компонент `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="b301a-141">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="b301a-142">Этот код выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="b301a-142">This code does the following:</span></span>

- <span data-ttu-id="b301a-143">Ссылается на библиотеку React с помощью `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="b301a-143">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="b301a-144">Ссылается на компоненты Fabric (`PrimaryButton`, `IButtonProps`, `Label`), которые используются для создания `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="b301a-144">References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="b301a-145">Объявляет новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="b301a-145">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="b301a-146">Объявляет функцию `insertText` для обработки события кнопки `onClick`.</span><span class="sxs-lookup"><span data-stu-id="b301a-146">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="b301a-147">Определяет пользовательский интерфейс компонента React в функции `render`.</span><span class="sxs-lookup"><span data-stu-id="b301a-147">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="b301a-148">В HTML-разметке используются компоненты `Label` и `PrimaryButton` из Fabric React и указывается, что при подключения события `onClick` будет запускаться функция `insertText`.</span><span class="sxs-lookup"><span data-stu-id="b301a-148">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="b301a-149">Добавление компонента React в надстройку</span><span class="sxs-lookup"><span data-stu-id="b301a-149">Add the React component to your add-in</span></span>

<span data-ttu-id="b301a-150">Добавьте компонент `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\App.tsx** и выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b301a-150">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="b301a-151">Добавьте приведенный ниже оператор импорта для ссылки на `ButtonPrimaryExample` из **Button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="b301a-151">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="b301a-152">Удалите два приведенные ниже оператора импорта.</span><span class="sxs-lookup"><span data-stu-id="b301a-152">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="b301a-153">Замените функцию по умолчанию `render()` на приведенный ниже код, в котором используется `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="b301a-153">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

  4. <span data-ttu-id="b301a-154">Сохраните изменения, внесенные в **App.tsx**.</span><span class="sxs-lookup"><span data-stu-id="b301a-154">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="b301a-155">Результат</span><span class="sxs-lookup"><span data-stu-id="b301a-155">See the result</span></span>

<span data-ttu-id="b301a-156">После сохранения изменений в **App.tsx** область задач надстройки в Word обновляется автоматически. </span><span class="sxs-lookup"><span data-stu-id="b301a-156">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="b301a-157">Текст по умолчанию и кнопка в нижней части области задач теперь отображают пользовательский интерфейс, определяемый компонентом `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="b301a-157">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="b301a-158">Нажмите кнопку **Вставить текст...** для вставки текста в документ.</span><span class="sxs-lookup"><span data-stu-id="b301a-158">Choose the **Insert text...** button to insert text into the document.</span></span>

![Снимок экрана c приложением Word с выделенными кнопкой "Вставить текст..." и предшествующим текстом](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="b301a-160">Поздравляем! Вы успешно создали надстройку области задач с помощью React и Office UI Fabric React!</span><span class="sxs-lookup"><span data-stu-id="b301a-160">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span> 

## <a name="see-also"></a><span data-ttu-id="b301a-161">См. также</span><span class="sxs-lookup"><span data-stu-id="b301a-161">See also</span></span>

- [<span data-ttu-id="b301a-162">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="b301a-162">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="b301a-163">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="b301a-163">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="b301a-164">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b301a-164">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="b301a-165">Начало работы с примером кода Fabric React</span><span class="sxs-lookup"><span data-stu-id="b301a-165">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
