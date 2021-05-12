---
title: Fluent UI React в Office надстройки
description: Узнайте, как использовать интерфейс Fluent React в Office надстройки.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cb7f04c21a52a2e4a3f271abc56aa325dd2b02fd
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330145"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>Использование интерфейса Fluent React в Office надстройки

Fluent UI React является официальной интерфейсной платформой JavaScript с открытым исходным кодом, предназначенной для создания интерфейсных интерфейсов, которые легко вписываются в широкий спектр продуктов Майкрософт, включая Office. Он предоставляет надежные, доступные React компоненты с высокой настройкой с помощью CSS-in-JS.

> [!NOTE]
> В этой статье описывается использование интерфейса fluent React в контексте Office надстройки. Но он также используется в широком диапазоне Microsoft 365 приложений и расширений. Дополнительные сведения см. в [странице Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) и репо с открытым исходным кодом [Fluent UI Web.](https://github.com/microsoft/fluentui)

В этой статье описывается, как создать надстройку, созданную с React и использующую компоненты React пользовательского интерфейса Fluent.

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

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки](../images/yo-office-word-react.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Выполните указанные ниже действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman. Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.

    > [!TIP]
    > Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду. После выполнения этой команды запустится локальный веб-сервер.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Чтобы проверить надстройку в Word, выполните приведенную ниже команду в корневом каталоге своего проекта. При этом запускается локальный веб-сервер (если он еще не запущен) и открывается приложение Word с загруженной надстройкой.

        ```command&nbsp;line
        npm start
        ```

    - Чтобы проверить надстройку в Word в браузере, выполните приведенную ниже команду в корневом каталоге проекта. После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).

        ```command&nbsp;line
        npm run start:web
        ```

        Чтобы использовать надстройку, откройте новый документ в Word в Интернете, а затем загрузите неопубликованную надстройку, следуя инструкциям из статьи [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

3. Чтобы открыть области задач надстройки, на вкладке **Главная** выберите кнопку **Показать задачу.** Обратите внимание на текст по умолчанию и кнопку **Запустить** в нижней части области задач. В остальной части этого поголовия вы переопределяете этот текст и кнопку, создав компонент React, использующий компоненты UX из пользовательского интерфейса Fluent React.

    ![Снимок экрана, показывающий приложение Word с выделенной кнопкой ленты Show Taskpane и кнопкой Run и сразу же предшествующим текстом, выделенным в области задач.](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>Создайте компонент React, использующий интерфейс Fluent React

На этом этапе вы уже создали самую простую надстройку в области задач c использованием React. Теперь выполните приведенные ниже действия, чтобы создать новый компонент React (`ButtonPrimaryExample`) в проекте надстройки. Компонент использует `Label` компоненты и `PrimaryButton` компоненты из пользовательского интерфейса Fluent React.

1. Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\taskpane\components**.
2. Создайте в этой папке новый файл под названием **Button.tsx**.
3. Введите в файл **Button.tsx** приведенный ниже код, чтобы определить компонент `ButtonPrimaryExample`.

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

Этот код выполняет следующие действия:

- Ссылается на библиотеку React с помощью `import * as React from 'react';`.
- Ссылается на компоненты React пользовательского интерфейса Fluent , `PrimaryButton` которые используются для создания `IButtonProps` `Label` `ButtonPrimaryExample` .
- Объявляет новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.
- Объявляет функцию `insertText` для обработки события кнопки `onClick`.
- Определяет пользовательский интерфейс компонента React в функции `render`. Разметка HTML использует и компоненты из интерфейса Fluent React и указывает, что при запуске события функция `Label` `PrimaryButton` будет `onClick` `insertText` работать.

## <a name="add-the-react-component-to-your-add-in"></a>Добавление компонента React в надстройку

Добавьте компонент `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\App.tsx** и выполните указанные ниже действия.

1. Добавьте приведенный ниже оператор импорта для ссылки на `ButtonPrimaryExample` из **Button.tsx**.

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. Удалите два приведенные ниже оператора импорта.

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
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

![Снимок экрана, показывающий приложение Word с текстом "Вставить текст...". кнопка и сразу перед выделением текста](../images/word-task-pane-with-react-component.png)

Поздравляем, вы успешно создали надстройку области задач с React и fluent UI React!

## <a name="see-also"></a>См. также

- [Word Add-in GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Fabric Core в Office надстройки](fabric-core.md)
- [Конструктивные шаблоны для надстроек Office](ux-design-pattern-templates.md)
