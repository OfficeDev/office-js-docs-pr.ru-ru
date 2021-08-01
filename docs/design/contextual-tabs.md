---
title: Создание настраиваемой контекстной вкладки Office надстроек
description: Узнайте, как добавить настраиваемые контекстные вкладки в Office надстройку.
ms.date: 07/15/2021
localization_priority: Normal
ms.openlocfilehash: 8696a9a7815b39ddd0100b70f7f9eaa94b1f4a89
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671536"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Создание настраиваемой контекстной вкладки Office надстроек

Контекстная вкладка — это скрытый контроль вкладок в ленте Office, отображаемой в строке вкладок, когда указанное событие происходит в Office документе. Например, **вкладка "Дизайн** таблицы", которая отображается на Excel при выборе таблицы. Вы включаете настраиваемые контекстные вкладки в Office надстройки и указываете, когда они видны или скрыты, создав обработчики событий, которые изменяют видимость. (Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

[!INCLUDE [Animation of contextual tabs and enabling buttons](../includes/animation-contextual-tabs-enable-button.md)]

> [!IMPORTANT]
> Пользовательские контекстные вкладки в настоящее время поддерживаются только на Excel и только на этих платформах и сборках:
>
> - Excel на Windows (только Microsoft 365 подписка): Версия 2102 (сборка 13801.20294) или более поздней версии.
> - Excel в Интернете

> [!NOTE]
> Настраиваемые контекстные вкладки работают только на платформах, поддерживаюх следующие наборы требований. Дополнительные подробности о наборах требований и работе с ними см. в Office [приложений и API.](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Вы можете использовать проверки времени запуска в коде, чтобы проверить, поддерживает ли комбинация хост и платформа пользователя эти наборы требований, описанные в описании Office приложений и [требований API.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) (Метод указания наборов требований в манифесте, который также описан в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, вы можете [реализовать альтернативный интерфейс интерфейса, если пользовательские](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)контекстные вкладки не поддерживаются.

## <a name="behavior-of-custom-contextual-tabs"></a>Поведение пользовательских контекстных вкладок

Пользовательский интерфейс для пользовательских контекстных вкладок следует шаблону встроенных Office контекстных вкладок. Ниже приводится базовый принцип размещения пользовательских контекстных вкладок.

- Когда отображается настраиваемая контекстная вкладка, она отображается на правом конце ленты.
- Если одна или несколько встроенных контекстных вкладок и одна или несколько пользовательских контекстных вкладок из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.
- Если надстройка имеет несколько контекстных вкладок и есть контексты, в которых видно несколько, они отображаются в порядке, в котором они определены в вашей надстройке. (Это направление в том же направлении, что и язык Office, то есть слева направо на левом и правом языках, но справа налево на языках справа налево.) Сведения [о том,](#define-the-groups-and-controls-that-appear-on-the-tab) как их определить, см. в материале Определение групп и элементов управления, которые отображаются на вкладке.
- Если несколько надстроек имеет контекстную вкладку, которая видна в определенном контексте, они отображаются в порядке запуска надстроек.
- Настраиваемые *контекстные* вкладки, в отличие от настраиваемой основной вкладки, не добавляются Office ленту приложения. Они присутствуют только в Office документах, на которых работает надстройка.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Основные действия по включаемой контекстной вкладке в надстройку

Ниже приводится основные действия для добавления настраиваемой контекстной вкладки в надстройку.

1. Настройте надстройку для использования общего времени запуска.
1. Определите вкладку, группы и элементы управления, которые отображаются на ней.
1. Зарегистрируйте контекстную вкладку с помощью Office.
1. Укажите обстоятельства, когда вкладка будет видна.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Добавление настраиваемой контекстной вкладки требует от надстройки использовать общее время работы. Дополнительные сведения см. в [раздел Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Определение групп и элементов управления, которые отображаются на вкладке

В отличие от настраиваемой вкладки ядра, которые определяются с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время запуска с помощью BLOB JSON. Код разрезает blob в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Настраиваемые контекстные вкладки присутствуют только в документах, на которых в настоящее время запущена надстройка. Это отличается от настраиваемой основной вкладки, которые добавляются в ленту Office приложения при установке надстройки и остаются в момент открытия другого документа. Кроме того, `requestCreateControls` метод может запускаться только один раз в сеансе надстройки. Если он снова вызван, ошибка будет выброшена.

> [!NOTE]
> Структура свойств и свойств BLOB JSON (и имен ключей) примерно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его элементов потомка в манифесте XML.

Мы пошаговую соберем пример контекстных вкладок JSON blob. Полная схема контекстной вкладки JSON находится [dynamic-ribbon.schema.js.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json) Если вы работаете в Visual Studio Code, вы можете использовать этот файл для получения IntelliSense и проверки JSON. Дополнительные сведения см. в [статью Редактирование JSON с Visual Studio Code - схемы и параметры JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Начните с создания строки JSON с двумя свойствами массива с `actions` именем и `tabs` . Массив — это спецификация всех функций, которые можно выполнять с помощью `actions` элементов управления на контекстной вкладке. Массив определяет одну или несколько контекстных вкладок, не более `tabs` *20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Этот простой пример контекстной вкладки будет иметь только одну кнопку и, следовательно, только одно действие. Добавьте следующее как единственный член `actions` массива. Об этой разметки обратите внимание:

    - Свойства `id` `type` и свойства обязательны.
    - Значение может `type` быть "ExecuteFunction" или "ShowTaskpane".
    - Свойство `functionName` используется только при значении `type` `ExecuteFunction` . Это имя функции, определенной в FunctionFile. Дополнительные сведения о FunctionFile см. в базовых [понятиях команд надстройки.](add-in-commands.md)
    - На более позднем этапе вы соберете это действие на кнопку на вкладке contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Добавьте следующее как единственный член `tabs` массива. Об этой разметки обратите внимание:

    - Свойство `id` является обязательным. Используйте краткий описательный ID, уникальный среди всех контекстных вкладок в надстройке.
    - Свойство `label` является обязательным. Это удобное строка, которая служит меткой контекстной вкладки.
    - Свойство `groups` является обязательным. Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь по крайней мере один член *и не более 20*. (Существует также ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, что также ограничивает количество групп, которые у вас есть. Дополнительные сведения см. в следующем шаге.)

    > [!NOTE]
    > Объект вкладки также может иметь необязательное свойство, которое указывает, видна ли вкладка сразу после `visible` начала надстройки. Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не вызовет их видимость (например, если пользователь выбирает объект определенного типа в документе), свойство по умолчанию не будет `visible` `false` присутствовать. В более позднем разделе мы покажем, как настроить свойство в ответ `true` на событие.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. В простом непрерывном примере контекстная вкладка имеет только одну группу. Добавьте следующее как единственный член `groups` массива. Об этой разметки обратите внимание:

    - Все свойства необходимы.
    - Свойство должно быть уникальным среди всех групп на `id` вкладке. Используйте краткий описательный ID.
    - Строка является удобной для `label` пользователя, которая служит в качестве метки группы.
    - Значение свойства — массив объектов, которые указывают значки, которые будут иметься у группы на ленте в зависимости от размера ленты и `icon` окна Office приложения.
    - Значение свойства — это массив объектов, которые указывают кнопки и `controls` меню в группе. Должно быть по крайней мере одно.

    > [!IMPORTANT]
    > *Общее число элементов управления на всей вкладке может быть не более 20.* Например, можно иметь 3 группы с 6 элементами управления и четвертую группу с 2 элементами управления, но нельзя иметь 4 группы с 6 элементами управления каждой.  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. Каждая группа должна иметь значок не менее двух размеров: 32x32 px и 80x80 px. Кроме того, можно использовать значки размеров 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px. Office определяет, какой значок использовать в зависимости от размера ленты и Office окна приложения. Добавьте следующие объекты в массив значок. (Если размеры окна и ленты достаточно большие для  появления хотя бы одного из элементов управления в группе, то не отображается значок группы. Например, просмотрите группу **Стилей** на ленте Word при сжатии и расширении окна Word.) Об этой разметки обратите внимание:

    - Необходимы оба свойства.
    - Единица `size` свойства измерения — пиксели. Значки всегда квадратные, поэтому число — это как высота, так и ширина.
    - Свойство `sourceLocation` указывает полный URL-адрес значка.

    > [!IMPORTANT]
    > Как правило, при переходе от разработки к производству (например, при изменении домена с локального на contoso.com) необходимо изменить URL-адреса в контекстных вкладок JSON.

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. В нашем простом непрерывном примере у группы есть только одна кнопка. Добавьте следующий объект как единственный член `controls` массива. Об этой разметки обратите внимание:

    - Все свойства, за `enabled` исключением, необходимы.
    - `type` указывает тип управления. Значения могут быть "Button", "Menu" или "MobileButton".
    - `id` может быть до 125 символов. 
    - `actionId` должен быть ID действия, определенного в `actions` массиве. (См. шаг 1 этого раздела.)
    - `label` является удобной строкой, которая служит в качестве метки кнопки.
    - `superTip` представляет собой богатую форму подсказки инструмента. Требуются `title` `description` как свойства, так и свойства.
    - `icon` указывает значки для кнопки. Предыдущие замечания о значке группы применяются и здесь.
    - `enabled` (необязательный) указывает, включена ли кнопка при запусках контекстной вкладки. Если по умолчанию `true` нет. 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
Ниже приводится полный пример BLOB JSON.

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Регистрация контекстной вкладки с помощью Office с помощью requestCreateControls

Контекстная вкладка регистрируется с помощью Office путем вызова [метода Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Обычно это делается в функции, назначенной или `Office.initialize` с помощью `Office.onReady` метода. Дополнительные данные об этих методах и инициализации надстройки см. в Office [надстройки.](../develop/initialize-add-in.md) Однако вы можете вызвать метод в любое время после инициализации.

> [!IMPORTANT]
> Метод может быть вызван только один раз в `requestCreateControls` заданном сеансе надстройки. Ошибка будет выброшена, если она будет вызвана снова.

Ниже приведен пример. Обратите внимание, что строка JSON должна быть преобразована в объект JavaScript с помощью метода, прежде чем она может быть передана `JSON.parse` функции JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Укажите контексты, когда вкладка будет видна с помощью requestUpdate

Как правило, настраиваемая контекстная вкладка должна отображаться, когда инициированное пользователем событие меняет контекст надстройки. Рассмотрим сценарий, в котором вкладка должна быть видна при активации диаграммы (по умолчанию в Excel книге).

Начните с назначения обработчиков. Обычно это делается в методе, как в следующем примере, который назначает обработчики (созданные на более позднем этапе) к событиям и событиям всех диаграмм в `Office.onReady` `onActivated` `onDeactivated` таблице.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Далее определите обработчики. Ниже приводится простой пример ошибки `showDataTab` [HostRestartNeeded,](#handle-the-hostrestartneeded-error) но см. ниже в этой статье для более надежной версии функции. Вот что нужно знать об этом коде:

- Office определяет время обновления состояния ленты. Метод [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestUpdate_input_) очереди запроса на обновление. Метод разрешит объект сразу после очереди запроса, а не после обновления `Promise` ленты.
- Параметром метода является объект `requestUpdate` [RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) указывает вкладку по своему ID точно так, как указано в *JSON* и (2) указывает видимость вкладки.
- Если у вас есть несколько пользовательских контекстных вкладок, которые должны быть видны в том же контексте, вы просто добавляете дополнительные объекты вкладок в `tabs` массив.

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

Обработник для сокрытия вкладки почти идентичен, за исключением того, что он задает `visible` свойство обратно `false` .

Библиотека Office JavaScript также предоставляет несколько интерфейсов (типов), чтобы упростить построение `RibbonUpdateData` объекта. Ниже приводится `showDataTab` функция TypeScript, которая использует эти типы.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Обзор вкладок и состояние включенной кнопки одновременно

Метод также используется для настройки включенного или отключенного состояния настраиваемой кнопки на настраиваемой контекстной вкладке или настраиваемой основной `requestUpdate` вкладке. Дополнительные сведения см. в материале [Enable and Disable Add-in Commands.](disable-add-in-commands.md) Возможны сценарии, в которых одновременно необходимо изменить видимость вкладки и состояние включенной кнопки. Вы делаете это одним вызовом `requestUpdate` . Ниже приводится пример, в котором кнопка на основной вкладке включена одновременно с тем, как отображается контекстная вкладка.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

В следующем примере включенная кнопка находится на той же контекстной вкладке, которая делается видимой.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="open-a-task-pane-from-contextual-tabs"></a>Откройте области задач из контекстных вкладок

Чтобы открыть области задач с кнопки на настраиваемой контекстной вкладке, создайте действие в JSON с `type` помощью `ShowTaskpane` . Затем определите кнопку с `actionId` набором свойств `id` к действию. Это открывает области задач по умолчанию, указанные `<Runtime>` элементом в манифесте.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

Чтобы открыть любую области задач, которая не является области задач по умолчанию, укажите свойство в `sourceLocation` определении действия. В следующем примере с другой кнопки открывается вторая области задач.

> [!IMPORTANT]
>
> - Если `sourceLocation` для действия задана задача, то  в области задач не используется общее время запуска. Он выполняется в новом времени запуска JavaScript.
> - Не более одной области задач может использовать совместное время работы, поэтому не более одного действия типа могут `ShowTaskpane` опустить `sourceLocation` свойство.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>Локализация текста JSON

BLOB JSON, который передается, не локализован так же, как локализована разметка манифеста для настраиваемой вкладки ядра (которая описывается при локализации Control из `requestCreateControls` [манифеста).](../develop/localization.md#control-localization-from-the-manifest) Вместо этого локализация должна происходить во время запуска с использованием отдельных BLOB-меток JSON для каждого локального. Мы рекомендуем использовать заявление, которое проверяет `switch` [свойство Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Ниже приведен пример.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

Затем код вызывает функцию, чтобы получить локализованный blob, который `requestCreateControls` передается, как в следующем примере.

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Лучшие практики для настраиваемой контекстной вкладки

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Реализация альтернативного интерфейса, когда пользовательские контекстные вкладки не поддерживаются

Некоторые сочетания платформы, Office приложения и Office сборки не `requestCreateControls` поддерживаются. Надстройка должна быть разработана для предоставления альтернативного опыта пользователям, которые запускают надстройки в одной из этих комбинаций. В следующих разделах описаны два способа предоставления впечатления от отката.

#### <a name="use-noncontextual-tabs-or-controls"></a>Использование неконтекстуальных вкладок или элементов управления

Существует элемент манифеста [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)который предназначен для создания впечатления от отката в надстройке, которая реализует настраиваемые контекстные вкладки, когда надстройка запущена на приложении или платформе, которая не поддерживает настраиваемые контекстные вкладки. 

Простейшая стратегия использования этого элемента заключается в *том,* что вы определяете в манифесте одну или несколько настраиваемых вкладки ядра (то есть неконтекстуальные пользовательские вкладки), дублирующие настройки ленты пользовательских контекстных вкладок в надстройке. Но вы `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` добавляете в качестве первого детского элемента [CustomTab](../reference/manifest/customtab.md). Эффект от этого ниже:

- Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то настраиваемая вкладка ядра не будет отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка запускается на  приложении или платформе, которые не поддерживаются, на ленте появится настраиваемая вкладка `requestCreateControls` ядра.

Ниже приводится пример этой простой стратегии.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

Эта простая стратегия использует настраиваемую вкладку ядра, которая зеркально отражает настраиваемую контекстную вкладку с ее детскими группами и средствами управления, но можно использовать более сложную стратегию. Элемент также может быть добавлен как (первый) детский элемент к элементам Group и Control (как тип кнопки, так и тип меню), а также `<OverriddenByRibbonApi>` элементам [](../reference/manifest/control.md#button-control) [](../reference/manifest/group.md) [](../reference/manifest/control.md) [](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` меню. Этот факт позволяет распространять группы и элементы управления, которые в противном случае отображаются на контекстной вкладке между различными группами, кнопками и меню в различных настраиваемой основной вкладке. Ниже приведен пример. Обратите внимание, что "MyButton" появится на настраиваемой вкладке ядра только в том случае, если пользовательские контекстные вкладки не поддерживаются. Но родительская группа и настраиваемая вкладка ядра будут отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Дополнительные примеры см. в [примере OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).

Если родительская вкладка, группа или меню помечены, то она не отображается, и все это детская разметка игнорируется, когда настраиваемые контекстные вкладки не `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` поддерживаются. Поэтому не имеет значения, имеет ли какой-либо из этих детских элементов элемент `<OverriddenByRibbonApi>` или его значение. Следствием этого является то, что если элемент меню, элемент управления или группа должны быть видны во всех контекстах, то не только он не должен быть отмечен, но и его предок меню, группа и вкладка также не должны быть отмечены таким образом `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . 

> [!IMPORTANT]
> Не *пометить* все детские элементы вкладки, группы или меню `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . Это бессмысленно, если родительский элемент помечен по причинам, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` заданным в предыдущем абзаце. Кроме того, если оставить на родительском (или установить его), то родитель будет отображаться независимо от того, поддерживаются ли пользовательские контекстные вкладки, но он будет пустым, когда они `<OverriddenByRibbonApi>` `false` поддерживаются. Таким образом, если все элементы ребенка не должны отображаться при поддержке настраиваемой контекстной вкладки, пометите родителя и только родителя с `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Использование API, которые показывают или скрывают области задач в указанных контекстах

В качестве альтернативы надстройке можно определить области задач с помощью элементов управления пользовательским интерфейсом, дублирующих функции элементов управления на настраиваемой `<OverriddenByRibbonApi>` контекстной вкладке. Затем используйте [методы Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) и [Office.addin.hide,](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) чтобы показать область задач, когда и только когда контекстная вкладка была бы показана, если она была поддержана. Дополнительные сведения об использовании этих методов см. в материале Показать или скрыть области задач [Office надстройки.](../develop/show-hide-add-in.md)

### <a name="handle-the-hostrestartneeded-error"></a>Обработка ошибки HostRestartNeeded

В некоторых случаях Office не может обновить ленту и возвращает ошибку. Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть. Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`. Код должен обрабатывать эту ошибку. Ниже приводится пример того, как. В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
