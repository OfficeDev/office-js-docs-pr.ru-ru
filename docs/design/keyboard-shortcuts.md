---
title: Настраиваемые клавиши в Office надстройки
description: Узнайте, как добавить в надстройку настраиваемые клавиши, также известные как комбинации ключей, Office надстройку.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0f4ef373ee5352f012561d76fa5bc01cb391af48
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151050"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Добавление настраиваемого сочетания клавиш в Office надстройки

Ярлыки клавиатуры, также известные как сочетания клавиш, позволяют пользователям надстройки работать более эффективно. Ярлыки клавиатуры также улучшают доступность надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Чтобы начать с рабочей версии надстройки с уже включенными клавишами, клонировать и запускать Excel [клавиши.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Если вы готовы добавить ярлыки клавиатуры в собственную надстройку, продолжи эту статью.

Существует три шага, чтобы добавить в надстройку ярлыки клавиатуры.

1. [Настройка манифеста надстройки.](#configure-the-manifest)
1. [Создание или изменение ярлыков JSON-файла для](#create-or-edit-the-shortcuts-json-file) определения действий и их клавиш.
1. [Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить функцию с каждым действием.

## <a name="configure-the-manifest"></a>Настройка манифеста

В манифест необходимо внести два небольших изменения. Один из них — включить надстройку для использования общего времени работы, а другой — указать на файл в формате JSON, в котором определены ярлыки клавиатуры.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общего времени работы

Добавление пользовательских ярлыков клавиатуры требует от надстройки использовать общее время работы. Дополнительные сведения: [Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

### <a name="link-the-mapping-file-to-the-manifest"></a>Привязка файла сопоставления к манифесту

Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Установите атрибут для полного URL-адреса файла JSON в проекте, который будет создан `Url` на более позднем этапе.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Создание или изменение ярлыков JSON-файла

Создайте файл JSON в проекте. Убедитесь, что путь файла соответствует расположению, указанному для атрибута элемента `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) В этом файле будут описаны ярлыки клавиатуры и действия, которые они будут вызывать.

1. В файле JSON есть два массива. Массив действий будет содержать объекты, которые определяют действия, которые будут вызываться, а массив ярлыков будет содержать объекты, которые соотносят комбинации ключей с действиями. Пример:
    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    Дополнительные сведения об объектах JSON см. в дополнительных сведениях [о конструкторе](#construct-the-action-objects) объектов действий [и создания объектов ярлыка.](#construct-the-shortcut-objects) Полная схема для ярлыков JSON находится в [расширенном манифесте.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > В этой статье можно использовать "CONTROL" на месте "Ctrl".

    На более позднем этапе действия сами будут соедему с функциями, которые вы пишете. В этом примере вы позже назовет SHOWTASKPANE функцией, которая вызывает метод, а HIDETASKPANE — функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Создание сопоставления действий с их функциями

1. В проекте откройте файл JavaScript, загруженный вашей htmL-страницей в `<FunctionFile>` элементе.
1. В файле JavaScript [используйте API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотнося каждое действие, указанное в файле JSON, с функцией JavaScript. Добавьте в файл следующий JavaScript. Обратите внимание на следующее о коде.

    - Первый параметр — это одно из действий из файла JSON.
    - Второй параметр — это функция, которая выполняется при нажатии клавиши на сочетание ключей, относясь к действию в файле JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.
1. Для тела функции используйте [метод Office.addin.showTaskpane](/javascript/api/office/office.addin#showAsTaskpane__) для открытия области задач надстройки. После этого код должен выглядеть следующим образом:

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. Добавьте второй вызов функции, чтобы соединить действие с функцией, вызываемой `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide.](/javascript/api/office/office.addin#hide__) Ниже приведен пример.

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

Следуя предыдущим шагам, надстройка позволяет управлять видимостью области задач, нажимая **на Ctrl+Alt+Up** и **Ctrl+Alt+Down.** Такое же поведение показано в примере Excel клавиш в репо Office PnP надстройки в GitHub. [](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)

## <a name="details-and-restrictions"></a>Сведения и ограничения

### <a name="construct-the-action-objects"></a>Построение объектов действия

При указании объектов в массиве `actions` ярлыков.json используйте следующие рекомендации.

- Имена свойств `id` и `name` обязательны.
- Свойство `id` используется для уникальной идентификации действия, вызываемого с помощью ярлыка клавиатуры.
- Свойство `name` должно быть удобной строкой, описываемой действием. Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".
- Свойство `type`— необязательное. В `ExecuteFunction` настоящее время поддерживается только тип.

Ниже приведен пример.

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

Полная схема для ярлыков JSON находится в [расширенном манифесте.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Построение объектов ярлыка

При указании объектов в массиве `shortcuts` ярлыков.json используйте следующие рекомендации.

- Имена свойств `action` `key` и `default` обязательно.
- Значение свойства является строкой и должно соответствовать одному из свойств `action` `id` объекта действия.
- Свойство может быть любым сочетанием символов A - Z, a -z, 0 - 9, а знаки препинания `default` "-", "_" и "+". (По соглашению в этих свойствах не используются буквы более низкого уровня.)
- Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (Alt, Ctrl, Shift) и только одного другого ключа.
- Shift не может использоваться в качестве только ключа модификатора. Объединяйте Shift с Alt или Ctrl.
- Для macs мы также поддерживаем ключ модификатора Команд.
- Для macs Alt соо- Для Windows командной командой нажата клавиша Ctrl.
- Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, Alt+a и Alt+A являются одним и тем же ярлыком, как `default` и Ctrl+- и Ctrl+, так как "-" и "_" являются одним и тем же физическим \_ ключом.
- Символ "+" указывает, что клавиши с обеих сторон нажаты одновременно.

Ниже приведен пример.

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

Полная схема для ярлыков JSON находится в [расширенном манифесте.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Клавиши KeyTips, также известные как последовательное клавиши, такие как ярлык Excel для выбора цвета заполнения **Alt+H, H,** не поддерживаются в Office надстроек.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Избегайте комбинаций ключей, которые используются другими надстройки

Существует множество клавиш, которые уже используются Office. Избегайте регистрации клавишных ярлыков для надстройки, которые уже используются, однако могут существовать некоторые случаи, когда необходимо переопределять существующие ярлыки клавиатуры или обрабатывать конфликты между несколькими надстройки, которые зарегистрировали один и тот же ярлык клавиатуры.

В случае конфликта пользователь увидит диалоговое окно при первой попытке использовать конфликтующий ярлык клавиатуры, обратите внимание, что имя действия, отображаемого в этом диалоговом диалоговом окне, является свойством в объекте действия в `name` `shortcuts.json` файле.

![Иллюстрация, показывающая конфликтный модал с двумя разными действиями для одного ярлыка.](../images/add-in-shortcut-conflict-modal.png)

Пользователь может выбрать, какое действие будет принимать ярлык клавиатуры. После выбора предпочтения сохраняются для будущих применений одного и того же ярлыка. Параметры ярлыка сохраняются для каждого пользователя, для платформы. Если пользователь хочет изменить свои предпочтения, он может **вызвать команду быстрого** доступа Office надстройки из поискового окна Tell **me.** При наводке команда очищает все параметры ярлыка надстройки пользователя, и пользователю снова будет предложен диалоговое окно конфликта при следующей попытке использовать конфликтующий ярлык.

![Поле поиска Tell me в Excel с указанием действия Office настройки ярлыка надстройки.](../images/add-in-reset-shortcuts-action.png)

Для наилучшего пользовательского интерфейса рекомендуется свести к минимуму конфликты с Excel с этими рекомендациями.

- Используйте только клавиши со следующим шаблоном: **Ctrl+Shift+Alt+* x****, где *x* — это другой ключ.
- Если вам нужно больше клавиш, ознакомьтесь со списком Excel клавиш [и](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f)не применяйте их в надстройки.
- Когда фокус клавиатуры находится внутри пользовательского интерфейса надстройки, **Ctrl+Spacebar** и **Ctrl+Shift+F10** не будут работать, так как это основные ярлыки доступности.
- На компьютере Windows или Mac, если в меню поиска недоступна команда "Reset Office надстройки", пользователь может вручную добавить команду в ленту, настроив ленту через контекстное меню.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Настройка ярлыков клавиатуры для платформы

Можно настроить ярлыки для конкретной платформы. Ниже приводится пример объекта, который настраивает ярлыки для каждой из следующих `shortcuts` платформ: `windows` , , `mac` `web` . Обратите внимание, что для каждого ярлыка необходимо иметь клавишу `default` ярлыка.

В следующем примере `default` ключом является клавиша отката для любой платформы, которая не указана. Единственная не указанная платформа Windows, поэтому ключ будет применяться только к `default` Windows.

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a>Локализовать ярлыки клавиатуры JSON

Если надстройка поддерживает несколько локалов, необходимо локализовать свойство `name` объектов действия. Кроме того, если в любом из локаутах, поддерживаюх надстройку, есть алфавиты или различные системы записи, а значит, и другие клавиатуры, возможно, потребуется также локализовать ярлыки. Сведения о том, как локализовать клавиши ярлыков JSON, см. в рубрезе [Localize extended overrides.](../develop/localization.md#localize-extended-overrides)

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Ярлыки браузера, которые нельзя переопределять

При использовании настраиваемого сочетания клавиш в Интернете некоторые клавиши, используемые браузером, не могут быть переопределены надстройки. Этот список находится в процессе выполнения. Если вы обнаружите другие комбинации, которые нельзя переопределять, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="next-steps"></a>Дальнейшие действия

- См. [Excel надстройки](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) для клавиатуры.
- Получите обзор работы с расширенными переопределениями в Работе с расширенными [переопределениями манифеста.](../develop/extended-overrides.md)
