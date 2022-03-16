---
title: Создание настраиваемой контекстной вкладки в Office надстроек
description: Узнайте, как добавить настраиваемые контекстные вкладки в Office надстройку.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 285a73b144470798e20d6d4ca374fb8a1655db2b
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511271"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Создание настраиваемой контекстной вкладки в Office надстроек

Контекстная вкладка — это скрытый контроль вкладок в ленте Office, отображаемой в строке вкладок, когда указанное событие происходит в Office документе. Например, **вкладка "Дизайн** таблицы", которая отображается на Excel при выборе таблицы. Вы включаете настраиваемые контекстные вкладки в Office надстройки и указываете, когда они видны или скрыты, создав обработчики событий, которые изменяют видимость. (Однако настраиваемые контекстные вкладки не реагируют на изменения фокуса.)

> [!NOTE]
> В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией. Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.
>
> - [Основные концепции команд надстроек](add-in-commands.md)

> [!IMPORTANT]
> Настраиваемые контекстные вкладки в настоящее время поддерживаются только на Excel и только на этих платформах и сборках.
>
> - Excel на Windows (только Microsoft 365 подписка): Версия 2102 (сборка 13801.20294) или более поздней версии.
> - Excel Mac: версия 16.53.806.0 или более поздней версии.
> - Excel в Интернете

> [!NOTE]
> Настраиваемые контекстные вкладки работают только на платформах, поддерживаюх следующие наборы требований. Дополнительные информацию о наборах требований и работе с ними см. в Office [приложений и API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Вы можете использовать проверки времени работы в коде, чтобы проверить, поддерживает ли комбинация хост и платформа пользователя эти наборы требований, описанные в проверках времени запуска для поддержки набора методов и [требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (Метод указания наборов требований в манифесте, который также описан в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, можно реализовать [альтернативный интерфейс,](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported) если пользовательские контекстные вкладки не поддерживаются.

## <a name="behavior-of-custom-contextual-tabs"></a>Поведение пользовательских контекстных вкладок

Пользовательский интерфейс пользовательских контекстных вкладок следует шаблону встроенных Office контекстных вкладок. Ниже приводится базовый принцип размещения пользовательских контекстных вкладок.

- Когда отображается настраиваемая контекстная вкладка, она отображается на правом конце ленты.
- Если одна или несколько встроенных контекстных вкладок и одна или несколько пользовательских контекстных вкладок из надстроек видны одновременно, настраиваемые контекстные вкладки всегда находятся справа от всех встроенных контекстных вкладок.
- Если надстройка имеет несколько контекстных вкладок и есть контексты, в которых видно несколько, они отображаются в порядке, в котором они определены в вашей надстройке. (Это направление в том же направлении, что и язык Office, то есть слева направо на языках слева направо, но справа налево на языках справа налево.) Сведения [о том,](#define-the-groups-and-controls-that-appear-on-the-tab) как их определить, см. в материале Определение групп и элементов управления, которые отображаются на вкладке.
- Если несколько надстроек имеет контекстную вкладку, которая видна в определенном контексте, они отображаются в порядке запуска надстроек.
- *Настраиваемые контекстные* вкладки, в отличие от настраиваемой основной вкладки, не добавляются Office ленте приложения. Они присутствуют только в Office документах, на которых работает надстройка.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Основные действия по включаемой контекстной вкладке в надстройку

Ниже приводится основные действия для добавления настраиваемой контекстной вкладки в надстройку.

1. Настройте надстройку для использования общего времени запуска.
1. Определите вкладку, группы и элементы управления, которые отображаются на ней.
1. Зарегистрируйте контекстную вкладку с помощью Office.
1. Укажите обстоятельства, когда вкладка будет видна.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Добавление настраиваемой контекстной вкладки требует от надстройки использовать общее время работы. Дополнительные сведения см. [в раздел Настройка надстройки для использования общего времени работы](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Определение групп и элементов управления, которые отображаются на вкладке

В отличие от настраиваемой вкладки ядра, которые определяются с помощью XML в манифесте, настраиваемые контекстные вкладки определяются во время запуска с помощью BLOB JSON. Код разрезает blob в объект JavaScript, а затем передает объект [методу Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)). Настраиваемые контекстные вкладки присутствуют только в документах, на которых в настоящее время запущена надстройка. Это отличается от настраиваемой основной вкладки, которые добавляются в ленту Office приложения при установке надстройки и остаются при открываемом другом документе. Кроме того, `requestCreateControls` метод может запускаться только один раз в сеансе надстройки. Если он снова вызван, ошибка будет выброшена.

> [!NOTE]
> Структура свойств и свойств BLOB JSON (и имен ключей) примерно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его элементов потомка в манифесте XML.

Мы пошаговую соберем пример контекстных вкладок JSON blob. Полная схема контекстной вкладки JSON находится на [динамической ленте.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Если вы работаете в Visual Studio Code, вы можете использовать этот файл для получения IntelliSense проверки JSON. Дополнительные сведения см. в Visual Studio Code [JSON с схемами и настройками JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Начните с создания строки JSON с двумя свойствами массива с именем `actions` и `tabs`. Массив `actions` — это спецификация всех функций, которые можно выполнять с помощью элементов управления на контекстной вкладке. Массив `tabs` определяет одну или несколько контекстных вкладок, не более *20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Этот простой пример контекстной вкладки будет иметь только одну кнопку и, следовательно, только одно действие. Добавьте следующее как единственный член массива `actions` . Об этой разметки обратите внимание:

    - Свойства `id` и `type` свойства обязательны.
    - Значение может быть `type` "ExecuteFunction" или "ShowTaskpane".
    - Свойство `functionName` используется только при значении `type` `ExecuteFunction`. Это имя функции, определенной в FunctionFile. Дополнительные сведения о FunctionFile см. в [базовых понятиях команд надстройки](add-in-commands.md).
    - На более позднем этапе вы соберете это действие на кнопку на вкладке contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Добавьте следующее как единственный член массива `tabs` . Об этой разметки обратите внимание:

    - Свойство `id` является обязательным. Используйте краткий описательный ID, уникальный среди всех контекстных вкладок в надстройке.
    - Свойство `label` является обязательным. Это удобное строка, которая служит меткой контекстной вкладки.
    - Свойство `groups` является обязательным. Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь не менее одного члена *и не более 20*. (Существует также ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, что также ограничивает количество групп, которые у вас есть. Дополнительные сведения см. в следующем шаге.)

    > [!NOTE]
    > Объект вкладки также может `visible` иметь необязательное свойство, которое указывает, видна ли вкладка сразу после начала надстройки. Так как контекстные вкладки обычно скрыты до тех пор, пока событие пользователя не вызовет их видимость (например, если пользователь выбирает объект определенного типа в документе), `visible` `false` свойство по умолчанию не будет присутствовать. В более позднем разделе мы покажем, как `true` настроить свойство в ответ на событие.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. В простом непрерывном примере контекстная вкладка имеет только одну группу. Добавьте следующее как единственный член массива `groups` . Об этой разметки обратите внимание:

    - Все свойства необходимы.
    - Свойство `id` должно быть уникальным среди всех групп манифеста. Используйте краткий описательный ID с 125 символами.
    - Строка `label` является удобной для пользователя, которая служит в качестве метки группы.
    - Значение `icon` свойства — массив объектов, которые указывают значки, которые будут иметься у группы на ленте в зависимости от размера ленты и окна Office приложения.
    - Значение `controls` свойства — это массив объектов, которые указывают кнопки и меню в группе. Должно быть по крайней мере одно.

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

1. Каждая группа должна иметь значок не менее двух размеров: 32x32 px и 80x80 px. Кроме того, можно использовать значки размеров 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px. Office определяет, какой значок использовать в зависимости от размера ленты и Office окна приложения. Добавьте следующие объекты в массив значок. (Если размеры окна и ленты достаточно большие для появления хотя бы одного из элементов  управления в группе, то не отображается значок группы. Например, просмотрите группу **Стилей** на ленте Word при сжатии и расширении окна Word.) Об этой разметки обратите внимание:

    - Необходимы оба свойства.
    - Единица `size` свойства измерения — пиксели. Значки всегда квадратные, поэтому число — это как высота, так и ширина.
    - Свойство `sourceLocation` указывает полный URL-адрес значка.

    > [!IMPORTANT]
    > Как правило, при переходе из разработки в производственную область (например, при изменении домена с локального на contoso.com) необходимо также изменить URL-адреса в контекстных вкладок JSON.

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

1. В нашем простом непрерывном примере у группы есть только одна кнопка. Добавьте следующий объект как единственный член массива `controls` . Об этой разметки обратите внимание:

    - Все свойства, за исключением `enabled`, необходимы.
    - `type` указывает тип управления. Значения могут быть "Button", "Menu" или "MobileButton".
    - `id` может быть до 125 символов.
    - `actionId` должен быть ID действия, определенного в массиве `actions` . (См. шаг 1 этого раздела.)
    - `label` является удобной строкой, которая служит в качестве метки кнопки.
    - `superTip` представляет собой богатую форму подсказки инструмента. Требуются `title` как `description` свойства, так и свойства.
    - `icon` указывает значки для кнопки. Предыдущие замечания о значке группы применяются и здесь.
    - `enabled` (необязательный) указывает, включена ли кнопка при запусках контекстной вкладки. Если по умолчанию нет.`true`

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

Контекстная вкладка регистрируется с помощью Office путем вызова [метода Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)). Обычно это делается в функции, `Office.initialize` назначенной или с помощью `Office.onReady` метода. Дополнительные данные об этих методах и инициализации надстройки см. в Office [надстройки](../develop/initialize-add-in.md). Однако вы можете вызвать метод в любое время после инициализации.

> [!IMPORTANT]
> Метод `requestCreateControls` может быть вызван только один раз в заданном сеансе надстройки. Ошибка будет выброшена, если она будет вызвана снова.

Ниже приведен пример. Обратите внимание, что строка JSON должна быть преобразована в объект JavaScript `JSON.parse` с помощью метода, прежде чем она может быть передана функции JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Укажите контексты, когда вкладка будет видна с помощью requestUpdate

Как правило, настраиваемая контекстная вкладка должна отображаться, когда инициированное пользователем событие меняет контекст надстройки. Рассмотрим сценарий, в котором вкладка должна быть видна при активации диаграммы (по умолчанию в Excel книге).

Начните с назначения обработчиков. Обычно это делается `Office.onReady` в методе, как в следующем примере, который назначает обработчики (созданные на более позднем этапе) `onActivated` `onDeactivated` к событиям и событиям всех диаграмм в таблице.

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

Далее определите обработчики. Ниже приводится простой `showDataTab`пример ошибки [HostRestartNeeded,](#handle-the-hostrestartneeded-error) но см. ниже в этой статье для более надежной версии функции. Вот что нужно знать об этом коде:

- Office определяет время обновления состояния ленты. Метод [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) очереди запроса на обновление. Метод разрешит объект `Promise` сразу после очереди запроса, а не после обновления ленты.
- Параметром `requestUpdate` метода является объект [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) , который (1) указывает вкладку по своему ID точно так, как указано в *JSON* и (2) указывает видимость вкладки.
- Если у вас есть несколько пользовательских контекстных вкладок, которые должны быть видны в том же контексте, вы просто добавляете дополнительные объекты вкладок в массив `tabs` .

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

Обработник для сокрытия вкладки почти идентичен, за исключением того, что он задает `visible` свойство обратно `false`.

Библиотека Office JavaScript также предоставляет несколько интерфейсов (типов),`RibbonUpdateData` чтобы упростить построение объекта. Ниже приводится функция `showDataTab` TypeScript, которая использует эти типы.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Обзор вкладок и состояние включенной кнопки одновременно

Метод `requestUpdate` также используется для настройки включенного или отключенного состояния настраиваемой кнопки на настраиваемой контекстной вкладке или настраиваемой основной вкладке. Дополнительные сведения см. в материале [Enable and Disable Add-in Commands](disable-add-in-commands.md). Возможны сценарии, в которых одновременно необходимо изменить видимость вкладки и состояние включенной кнопки. Вы делаете это одним вызовом `requestUpdate`. Ниже приводится пример, в котором кнопка на основной вкладке включена одновременно с тем, как отображается контекстная вкладка.

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

Чтобы открыть области задач с кнопки на настраиваемой контекстной вкладке, создайте действие в JSON с помощью `type` `ShowTaskpane`. Затем определите кнопку с `actionId` набором свойств к `id` действию. Это открывает области задач по умолчанию, заданные элементом **Runtime** в манифесте.

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

Чтобы открыть любую области задач, которая не является области задач по умолчанию, `sourceLocation` укажите свойство в определении действия. В следующем примере с другой кнопки открывается вторая области задач.

> [!IMPORTANT]
>
> - Если для `sourceLocation` действия задана задача, то в области задач *не используется общее* время запуска. Он выполняется в новом времени запуска JavaScript.
> - Не более одной области задач может использовать совместное время работы, `ShowTaskpane` поэтому не более одного действия типа могут опустить `sourceLocation` свойство.

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

BLOB JSON `requestCreateControls` , который передается, не локализован так же, как локализована разметка манифеста для пользовательских вкладок ядра (которая описана в локализацией [Control из манифеста](../develop/localization.md#control-localization-from-the-manifest)). Вместо этого локализация должна происходить во время запуска с использованием отдельных BLOB-меток JSON для каждого локального. Мы рекомендуем использовать заявление`switch`, которое проверяет [свойство Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member). Ниже приведен пример.

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

Затем код вызывает функцию, чтобы получить локализованный blob `requestCreateControls`, который передается, как в следующем примере.

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Лучшие практики для настраиваемой контекстной вкладки

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Реализация альтернативного интерфейса, когда пользовательские контекстные вкладки не поддерживаются

Некоторые сочетания платформы, Office приложения и Office сборки не поддерживаются`requestCreateControls`. Надстройка должна быть разработана для предоставления альтернативного опыта пользователям, которые запускают надстройки в одной из этих комбинаций. В следующих разделах описаны два способа предоставления впечатления от отката.

#### <a name="use-noncontextual-tabs-or-controls"></a>Использование неконтекстуальных вкладок или элементов управления

Существует элемент манифеста [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), который предназначен для создания впечатления от отката в надстройке, которая реализует настраиваемые контекстные вкладки при работе надстройки на приложении или платформе, не поддерживаюх настраиваемые контекстные вкладки.

Простейшая стратегия использования этого элемента заключается в том, чтобы определить одну или несколько настраиваемых вкладки ядра (  то есть неконтекстуальные пользовательские вкладки) в манифесте, дублирующем настройки ленты пользовательских контекстных вкладок в надстройке. Но вы добавляете `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` в качестве первого детского элемента элементы элементов ["Группа](../reference/manifest/group.md)[",](../reference/manifest/control.md) "Управление" и "Элемент меню" на настраиваемые вкладки ядра. Эффект от этого ниже:

- Если надстройка работает на приложении и платформе, поддерживаюх настраиваемые контекстные вкладки, то настраиваемые основные группы и элементы управления не будут отображаться на ленте. Вместо этого настраиваемая контекстная вкладка будет создана, когда надстройка вызывает `requestCreateControls` метод.
- Если надстройка работает на  `requestCreateControls`приложении или платформе, которые не поддерживаются, элементы отображаются на пользовательских вкладок ядра.

Ниже приведен пример. Обратите внимание, что "MyButton" появится на настраиваемой вкладке ядра только в том случае, если пользовательские контекстные вкладки не поддерживаются. Но родительская группа и настраиваемая вкладка ядра будут отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки.

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
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Дополнительные примеры см. в [примере OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).

Если родительская `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`группа или меню помечены, то она не отображается, и вся ее детская разметка игнорируется, когда настраиваемые контекстные вкладки не поддерживаются. Поэтому не важно, есть ли какой-либо из этих детских элементов элемент **OverriddenByRibbonApi** или его значение. Следствием этого является то, что если элемент меню или элемент управления должен быть виден во всех контекстах, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`то не только он не должен быть отмечен, но и его предок меню и группа также не должны быть отмечены *таким образом*.

> [!IMPORTANT]
> Не *пометить* все детские элементы группы или меню `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`с помощью . Это бессмысленно, если родительский элемент `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` помечен по причинам, заданным в предыдущем абзаце. Кроме того, если оставить **overriddenByRibbonApi** на родительском ( `false`или установить его), то родитель будет отображаться независимо от того, поддерживаются ли настраиваемые контекстные вкладки, но при поддержке они будут пустыми. Таким образом, если все элементы ребенка не должны отображаться при поддержке настраиваемой контекстной вкладки, пометите родительский элемент `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Использование API, которые показывают или скрывают области задач в указанных контекстах

В качестве альтернативы **OverriddenByRibbonApi** надстройка может определить области задач с помощью элементов управления пользовательским интерфейсом, дублирующих функции элементов управления на настраиваемой контекстной вкладке. Затем используйте [методы Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) и [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)), чтобы показать область задач, когда была бы показана контекстная вкладка при ее поддержке. Дополнительные сведения об использовании этих методов см. в материале Показать или скрыть области задач [Office надстройки](../develop/show-hide-add-in.md).

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

## <a name="resources"></a>Ресурсы

- [Пример кода: создание настраиваемой контекстной вкладки на ленте](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Community пример контекстных вкладок

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]