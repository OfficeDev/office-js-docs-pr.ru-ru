---
title: Использование Office UI Fabric React в надстройках Office
description: ''
ms.date: 12/04/2017
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>Использование Office UI Fabric React в надстройках Office

Office UI Fabric — это интерфейсная платформа JavaScript для построения взаимодействия с пользователем в Office и Office 365. Если вы разрабатываете надстройку с использованием React, пользовательский интерфейс рекомендуется создать с помощью Fabric React. В Fabric предоставлены некоторые компоненты дизайна на основе React, например кнопки и флажки, которые можно использовать в надстройке.

Чтобы использовать компоненты Fabric React в своей надстройке, выполните указанные ниже действия.

> [!NOTE]
> Если вы выполните действия, описанные в этой статье, в надстройке также будет доступен компонент Fabric Core.

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>Шаг 1. Создание проекта с помощью генератора Yeoman для Office

Чтобы создать надстройку, в которой используется Fabric React, рекомендуется использовать генератор Yeoman для Office. Генератор Yeoman для Office обеспечивает формирование шаблонов для проектов и управление сборкой, необходимые для разработки надстройки Office.

Чтобы создать проект, выполните следующие действия, используя **Windows PowerShell** (а не командную строку):

1. Установите необходимые компоненты.
2. Запустите `yo office`, чтобы создать файлы проекта для надстройки.
3. Когда вам будет предложено выбрать клиентское приложение Office, выберите **Word**.
4. Перейдите к каталогу с файлами проекта и запустите `npm start`. Автоматически откроется окно браузера с вертушкой.
5. [Загрузите неопубликованный манифест](..\testing\test-debug-office-add-ins.md), чтобы просмотреть весь пользовательский интерфейс надстройки.

## <a name="step-2---add-a-fabric-react-component"></a>Шаг 2. Добавление компонента Fabric React

Теперь добавьте в надстройку компоненты Fabric React. Создайте компонент React под названием `ButtonPrimaryExample`, который состоит из элементов Label и PrimaryButton из Fabric React. Создание `ButtonPrimaryExample`

1. Откройте папку проекта, созданную генератором Yeoman, и перейдите в раздел **src\components**.
2. Создайте файл **button.tsx**.
3. В файле **button.tsx** введите указанный код, чтобы создать компонент `ButtonPrimaryExample`.

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

Этот код выполняет следующие действия:

- Ссылается на библиотеку React с помощью `import * as React from 'react';`.
- Ссылается на компоненты Fabric (PrimaryButton, IButtonProps, Label), которые используются для создания `ButtonPrimaryExample`.
- Объявляет и публикует новый компонент `ButtonPrimaryExample` с помощью `export class ButtonPrimaryExample extends React.Component`.
- Объявляет функцию `insertText` для обработки события `onClick`.
- Определяет пользовательский интерфейс компонента React в функции `render`. Отрисовка определяет структуру компонента. В `render` для подключения события `onClick` используется `this.insertText`.

## <a name="step-3---add-the-react-component-to-your-add-in"></a>Шаг 3. Добавление компонента React в надстройку

Добавьте `ButtonPrimaryExample` к своей надстройке. Для этого откройте файл **src\components\app.tsx** и выполните перечисленные действия.

- Добавьте указанный оператор импорта для ссылки на `ButtonPrimaryExample` из файла **button.tsx**, созданного в шаге 2 (расширение файла не требуется).

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- Замените функцию `render()` по умолчанию на приведенный ниже код, в котором используется `<ButtonPrimaryExample />`.

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

Сохраните изменения. Все открытые экземпляры браузеров, включая надстройку, автоматически обновятся и отобразят компонент React `ButtonPrimaryExample`. Обратите внимание, что текст по умолчанию и кнопка заменяются текстом и основной кнопкой, определенной в `ButtonPrimaryExample`.

## <a name="recommended-components"></a>Рекомендуемые компоненты

Ниже приведен список компонентов дизайна Fabric React, которые мы рекомендуем использовать в надстройке.

- [Строка навигации](breadcrumb.md)
- [Кнопка](button.md)
- [Флажок](checkbox.md)
- [ChoiceGroup](choicegroup.md)
- [Раскрывающееся меню](dropdown.md)
- [Подпись](label.md)
- [Список](list.md)
- [Сводка](pivot.md)
- [TextField](textfield.md)
- [Переключатель](toggle.md)

> [!NOTE]
> Со временем мы добавим другие компоненты.

## <a name="see-also"></a>См. также

- [Office UI Fabric React](https://dev.office.com/fabric#/)
- [Начало работы с примером кода Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Конструктивные шаблоны (используется Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Пример пользовательского интерфейса Fabric для надстройки Office (используется Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Использование Fabric 2.6.1 в надстройке Office](ui-elements/using-office-ui-fabric.md)
- [Генератор Yeoman для Office](https://github.com/OfficeDev/generator-office)
