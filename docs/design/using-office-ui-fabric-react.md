---
title: Использование Office UI Fabric React в надстройках Office
description: ''
ms.date: 2/28/2019
localization_priority: Priority
ms.openlocfilehash: 7d3e280298ee6761be9e7ced96d3490defeef7f0
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359242"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="ad30a-102">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="ad30a-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="ad30a-p101">Office UI Fabric — это интерфейсная платформа JavaScript для построения взаимодействия с пользователем в Office и Office 365. Если вы разрабатываете надстройку с использованием React, пользовательский интерфейс рекомендуется создать с помощью Fabric React. В Fabric предоставлены некоторые компоненты дизайна на основе React, например кнопки и флажки, которые можно использовать в надстройке.</span><span class="sxs-lookup"><span data-stu-id="ad30a-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="ad30a-106">Чтобы использовать компоненты Fabric React в своей надстройке, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ad30a-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="ad30a-107">Если вы выполните действия, описанные в этой статье, в надстройке также будет доступен компонент Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="ad30a-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="ad30a-108">Шаг 1. Создание проекта с помощью генератора Yeoman для Office</span><span class="sxs-lookup"><span data-stu-id="ad30a-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="ad30a-109">Чтобы создать надстройку, в которой используется Fabric React, рекомендуется использовать генератор Yeoman для Office.</span><span class="sxs-lookup"><span data-stu-id="ad30a-109">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office.</span></span> <span data-ttu-id="ad30a-110">Генератор Yeoman для Office обеспечивает формирование шаблонов для проектов и управление сборкой, необходимые для разработки надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="ad30a-110">The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office Add-in.</span></span>

<span data-ttu-id="ad30a-111">Чтобы создать проект, выполните следующие действия, используя **Windows PowerShell** (а не командную строку):</span><span class="sxs-lookup"><span data-stu-id="ad30a-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="ad30a-112">Установите необходимые компоненты.</span><span class="sxs-lookup"><span data-stu-id="ad30a-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="ad30a-113">Запустите `yo office`, чтобы создать файлы проекта для надстройки.</span><span class="sxs-lookup"><span data-stu-id="ad30a-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="ad30a-114">Когда вам будет предложено выбрать клиентское приложение Office, выберите **Word**.</span><span class="sxs-lookup"><span data-stu-id="ad30a-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="ad30a-p103">Перейдите к каталогу с файлами проекта и запустите `npm start`. Автоматически откроется окно браузера с вертушкой.</span><span class="sxs-lookup"><span data-stu-id="ad30a-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="ad30a-117">[Загрузите неопубликованный манифест](..\testing\test-debug-office-add-ins.md), чтобы просмотреть весь пользовательский интерфейс надстройки.</span><span class="sxs-lookup"><span data-stu-id="ad30a-117">[Sideload your manifest](..\testing\test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="ad30a-118">Шаг 2. Добавление компонента Fabric React</span><span class="sxs-lookup"><span data-stu-id="ad30a-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="ad30a-p104">Теперь добавьте в надстройку компоненты Fabric React. Создайте компонент React под названием `ButtonPrimaryExample`, который состоит из элементов Label и PrimaryButton из Fabric React. Создание `ButtonPrimaryExample`</span><span class="sxs-lookup"><span data-stu-id="ad30a-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="ad30a-122">Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\components**.</span><span class="sxs-lookup"><span data-stu-id="ad30a-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="ad30a-123">Создайте файл **button.tsx**.</span><span class="sxs-lookup"><span data-stu-id="ad30a-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="ad30a-124">В файле **button.tsx** введите указанный код, чтобы создать компонент `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
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

<span data-ttu-id="ad30a-125">Этот код выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="ad30a-125">This code does the following:</span></span>

- <span data-ttu-id="ad30a-126">Ссылается на библиотеку React с помощью `import * as React from 'react';`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="ad30a-127">Ссылается на компоненты Fabric (PrimaryButton, IButtonProps, Label), которые используются для создания `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="ad30a-128">Объявляет и публикует новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="ad30a-129">Объявляет функцию `insertText` для обработки события `onClick`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="ad30a-p105">Определяет пользовательский интерфейс компонента React в функции `render`. Отрисовка определяет структуру компонента. В `render` для подключения события `onClick` используется `this.insertText`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="ad30a-133">Шаг 3. Добавление компонента React в надстройку</span><span class="sxs-lookup"><span data-stu-id="ad30a-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="ad30a-134">Добавьте `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\app.tsx** и выполните перечисленные действия.</span><span class="sxs-lookup"><span data-stu-id="ad30a-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="ad30a-135">Добавьте указанный оператор импорта для ссылки на `ButtonPrimaryExample` из файла **button.tsx**, созданного в шаге 2 (расширение файла не требуется).</span><span class="sxs-lookup"><span data-stu-id="ad30a-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="ad30a-136">Замените функцию `render()` по умолчанию на приведенный ниже код, в котором используется `<ButtonPrimaryExample />`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="ad30a-p106">Сохраните изменения. Все открытые экземпляры браузеров, включая надстройку, автоматически обновятся и отобразят компонент React `ButtonPrimaryExample`. Обратите внимание, что текст по умолчанию и кнопка заменяются текстом и основной кнопкой, определенной в `ButtonPrimaryExample`.</span><span class="sxs-lookup"><span data-stu-id="ad30a-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>



## <a name="see-also"></a><span data-ttu-id="ad30a-140">См. также</span><span class="sxs-lookup"><span data-stu-id="ad30a-140">See also</span></span>

- [<span data-ttu-id="ad30a-141">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="ad30a-141">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="ad30a-142">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ad30a-142">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="ad30a-143">Начало работы с примером кода Fabric React</span><span class="sxs-lookup"><span data-stu-id="ad30a-143">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="ad30a-144">Пример пользовательского интерфейса Fabric для надстройки Office (используется Fabric 1.0)</span><span class="sxs-lookup"><span data-stu-id="ad30a-144">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="ad30a-145">Генератор Yeoman для Office</span><span class="sxs-lookup"><span data-stu-id="ad30a-145">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
