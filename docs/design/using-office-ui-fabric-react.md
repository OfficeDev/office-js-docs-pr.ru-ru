---
title: Fluent UI React в надстройках Office
description: Узнайте, как использовать Fluent пользовательского React в Office надстройки.
ms.date: 04/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e9386af4a2930e686da857249029228472bcbc2
ms.sourcegitcommit: 3c5ede9c4f9782947cea07646764f76156504ff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/06/2022
ms.locfileid: "64682219"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>Использование Fluent пользовательского React в Office надстройки

Fluent пользовательского интерфейса React является официальной интерфейсной платформой JavaScript с открытым кодом, предназначенной для создания интерфейсов, которые легко помещаются в широкий спектр продуктов Майкрософт, включая Office. Он обеспечивает надежные, современные, доступные компоненты на основе React, которые легко настраиваются с помощью CSS-in-JS.

> [!NOTE]
> В этой статье описывается использование Fluent пользовательского React в контексте Office надстроек. Но он также используется в широком спектре Microsoft 365 и расширений. Дополнительные сведения см. [в Fluent пользовательского React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) и репозиторий открытый код [веб-интерфейса Fluent пользовательского интерфейса](https://github.com/microsoft/fluentui).

В этой статье описывается, как создать надстройку, созданную с React и использующую Fluent компоненты React пользовательского интерфейса.

## <a name="create-an-add-in-project"></a>Создание проекта надстройки

Чтобы создать надстройку с использованием React, рекомендуется воспользоваться генератором Yeoman для надстроек Office.

### <a name="install-the-prerequisites"></a>Установка необходимых компонентов

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a>Создание проекта

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project using React framework`
- **Выберите тип сценария:** `TypeScript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Word`

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-word-react.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Выполните следующие действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду. После выполнения этой команды запустится локальный веб-сервер.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Чтобы проверить надстройку в Word, выполните приведенную ниже команду в корневом каталоге своего проекта. При этом запускается локальный веб-сервер и открывается Word с загруженной надстройке.

        ```command&nbsp;line
        npm start
        ```

    - Чтобы проверить надстройку в Word в браузере, выполните приведенную ниже команду в корневом каталоге проекта. После выполнения этой команды запустится локальный веб-сервер. Замените "{url}" на URL-адрес документа Word в OneDrive или библиотеке SharePoint, для которой у вас есть разрешения.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

3. Чтобы открыть область задач надстройки, на вкладке **"** Главная" нажмите кнопку **"Показать область** задач". Обратите внимание на текст по умолчанию и кнопку **Запустить** в нижней части области задач. В оставшейся части этого пошагового руководства вы переопределите этот текст и кнопку, создав компонент React, использующий компоненты пользовательского интерфейса из Fluent пользовательского React.

    ![Снимок экрана: приложение Word с выделенной кнопкой ленты "Показать область задач", кнопкой "Выполнить" и текстом, выделенным непосредственно в области задач.](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>Создание компонента React, использующего Fluent пользовательского интерфейса React

На этом этапе вы уже создали самую простую надстройку в области задач c использованием React. Теперь выполните приведенные ниже действия, чтобы создать новый компонент React (`ButtonPrimaryExample`) в проекте надстройки. Компонент использует компоненты `Label` из `PrimaryButton` Fluent пользовательского React.

1. Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\taskpane\components**.
2. Создайте в этой папке новый файл под названием **Button.tsx**.
3. Введите в файл **Button.tsx** приведенный ниже код, чтобы определить компонент `ButtonPrimaryExample`.

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';

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

Этот код выполняет следующие действия:

- Ссылается на библиотеку React с помощью `import * as React from 'react';`.
- Ссылается Fluent пользовательского интерфейса React компонентов (`PrimaryButton`, , ), `Label`которые используются для создания`ButtonPrimaryExample``IButtonProps`.
- Объявляет новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.
- Объявляет функцию `insertText` для обработки события кнопки `onClick`.
- Определяет пользовательский интерфейс компонента React в функции `render`. В разметке HTML `Label` `PrimaryButton` `onClick` используются компоненты Fluent пользовательского интерфейса React и указывается, `insertText` что при срабатывании события будет выполняться функция.

## <a name="add-the-react-component-to-your-add-in"></a>Добавление компонента React в надстройку

Добавьте компонент `ButtonPrimaryExample` в надстройку, открыв **файл src\components\App.tsx** и выполнив следующие действия.

1. Добавьте приведенный ниже оператор импорта для ссылки на `ButtonPrimaryExample` из **Button.tsx**.

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. Удалите следующую инструкцию импорта.

    ```typescript
    import Progress from './Progress';
    ```

3. Замените функцию по умолчанию `render()` на приведенный ниже код, в котором используется `ButtonPrimaryExample`.

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

4. Сохраните изменения, внесенные в **App.tsx**.

## <a name="see-the-result"></a>Результат

После сохранения изменений в **App.tsx** область задач надстройки в Word обновляется автоматически.  Текст по умолчанию и кнопка в нижней части области задач теперь отображают пользовательский интерфейс, определяемый компонентом `ButtonPrimaryExample`. Нажмите кнопку **Вставить текст...** для вставки текста в документ.

![Снимок экрана: приложение Word с надписью "Вставка текста..." кнопка  и выделенный текст непосредственно перед текстом.](../images/word-task-pane-with-react-component.png)

Поздравляем, вы успешно создали надстройку области задач с помощью React и Fluent пользовательского React!

## <a name="see-also"></a>См. также

- [Надстройка Word GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Fabric Core в надстройках Office](fabric-core.md)
- [Конструктивные шаблоны для надстроек Office](ux-design-pattern-templates.md)
