---
title: Настраиваемые сочетания клавиш в надстройки Office
description: Узнайте, как добавить в надстройку Office пользовательские сочетания клавиш, также известные как сочетания клавиш.
ms.date: 11/22/2021
localization_priority: Normal
ms.openlocfilehash: 462e5bfdd4e7f825318d6affb631beafc7c08fe5
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423022"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Добавление настраиваемых сочетаний клавиш в надстройки Office

Сочетания клавиш, также называемые сочетаниями клавиш, позволяют пользователям надстройки работать эффективнее. Сочетания клавиш также улучшают специальные возможности надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Чтобы начать с рабочей версии надстройки с уже включенными сочетаниями клавиш, клонируйте и запустите примеры сочетаний [клавиш Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts). Когда вы будете готовы добавить сочетания клавиш в собственную надстройку, перейдите к этой статье.

Добавить сочетания клавиш в надстройку можно тремя шагами.

1. [Настройте манифест надстройки](#configure-the-manifest).
1. [Создайте или измените JSON-файл](#create-or-edit-the-shortcuts-json-file) ярлыков для определения действий и сочетаний клавиш.
1. [Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) среды выполнения API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) , чтобы сопоставить функцию с каждым действием.

## <a name="configure-the-manifest"></a>Настройка манифеста

В манифест необходимо внести два небольших изменения. Один из них — разрешить надстройке использовать общую среду выполнения, а другой — указать на файл в формате JSON, где вы определили сочетания клавиш.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Настройка надстройки для использования общей среды выполнения

Чтобы добавить настраиваемые сочетания клавиш, надстройка будет использовать [общую среду выполнения](../testing/runtimes.md#shared-runtime). Дополнительные сведения см. в разделе "Настройка [надстройки для использования общей среды выполнения"](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Связывание файла сопоставления с манифестом

Непосредственно *под* элементом манифеста (не внутри) **\<VersionOverrides\>** добавьте [элемент ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Задайте `Url` для атрибута полный URL-адрес JSON-файла в проекте, который будет создаваться на следующем шаге.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Создание или изменение JSON-файла ярлыков

Создайте JSON-файл в проекте. Убедитесь, что путь к файлу `Url` соответствует расположению, указанному для атрибута элемента [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . В этом файле описываются сочетания клавиш и действия, которые они будут вызывать.

1. Внутри JSON-файла есть два массива. Массив действий будет содержать объекты, определяющие вызываемые действия, а массив ярлыков будет содержать объекты, которые сопоставляют сочетания клавиш с действиями. Пример:
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

    Дополнительные сведения об объектах JSON см. в разделе ["Создание объектов действий](#construct-the-action-objects) и [создание объектов ярлыков"](#construct-the-shortcut-objects). Полная схема для ярлыков JSON находится в [файле extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > В этой статье вместо клавиши CTRL можно использовать control.

    На следующем шаге действия будут сопоставлены с функциями, которые вы пишете. В этом примере вы позже сопоставляете SHOWTASKPANE `Office.addin.showAsTaskpane` с функцией, которая вызывает метод, и HIDETASKPANE с функцией, которая вызывает `Office.addin.hide` метод.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Создание сопоставления действий с их функциями

1. В проекте откройте файл JavaScript, загруженный HTML-страницей в элементе **\<FunctionFile\>** .
1. В файле JavaScript используйте API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) , чтобы сопоставить каждое действие, указанное в JSON-файле, с функцией JavaScript. Добавьте в файл следующий код JavaScript. Обратите внимание на следующие сведения о коде.

    - Первый параметр — это одно из действий из JSON-файла.
    - Второй параметр — это функция, которая выполняется, когда пользователь нажимает сочетание клавиш, сопоставленное с действием в JSON-файле.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Чтобы продолжить пример, используйте его `'SHOWTASKPANE'` в качестве первого параметра.
1. В тексте функции используйте метод [Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) , чтобы открыть область задач надстройки. Когда все будет готово, код должен выглядеть следующим образом:

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

1. Добавьте второй вызов функции `Office.actions.associate` , чтобы сопоставить `HIDETASKPANE` действие с функцией, которая вызывает [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)). Ниже приведен пример.

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

После выполнения предыдущих действий надстройка может переключать видимость области задач, нажимая **клавиши CTRL+ALT+UP** и **CTRL+ALT+DOWN**. Такое же поведение показано в примере сочетаний клавиш [Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) в репозитории PnP надстроек Office в GitHub.

## <a name="details-and-restrictions"></a>Сведения и ограничения

### <a name="construct-the-action-objects"></a>Создание объектов действий

При указании объектов `actions` в массиве shortcuts.json следуйте приведенным ниже рекомендациям.

- Имена свойств являются `id` обязательными `name` .
- Свойство `id` используется для уникальной идентификации действия, вызываемого с помощью сочетания клавиш.
- Свойство `name` должно быть понятной строкой, описываемой действием. Это должно быть сочетание символов A — Z, a - z, 0 – 9 и знаков препинания "-", "_" и "+".
- Свойство `type`— необязательное. В настоящее `ExecuteFunction` время поддерживается только тип.

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

Полная схема для ярлыков JSON находится в [файле extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Создание объектов ярлыков

При указании объектов `shortcuts` в массиве shortcuts.json следуйте приведенным ниже рекомендациям.

- Имена свойств и `action``key``default` являются обязательными.
- Значение свойства является `action` строкой и должно соответствовать одному `id` из свойств в объекте действия.
- Свойство `default` может быть любым сочетанием символов A — Z, a -z, 0 – 9 и знаков препинания "-", "_" и "+". (По соглашению строчные буквы не используются в этих свойствах.)
- Свойство `default` должно содержать имя по крайней мере одного ключа модификатора (ALT, CTRL, SHIFT) и только одного другого ключа.
- Shift нельзя использовать в качестве только ключа модификатора. Объедините shift с помощью клавиш ALT или CTRL.
- Для компьютеров Mac также поддерживается ключ модификатора команд.
- Для компьютеров Mac ALT сопоставляется с ключом Option. Для Windows команда сопоставляется с клавишей CTRL.
- Если два символа связаны с одинаковым физическим ключом в стандартной клавиатуре, `default` они являются синонимами в свойстве. Например, ALT+A и ALT+A — это одно и то же сочетание клавиш, поэтому клавиши CTRL+- и CTRL+\_ , так как "-" и "_" являются одинаковыми физическими клавишами.
- Символ "+" указывает, что клавиши с любой стороны нажимаются одновременно.

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

Полная схема для ярлыков JSON находится в [файле extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Подсказки клавиш, также известные как последовательные сочетания клавиш, такие как ярлык Excel для выбора цвета заливки **ALT+H, H**, не поддерживаются в надстройки Office.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Избегайте сочетаний клавиш, используемых другими надстройки

Office уже использует множество сочетаний клавиш. Избегайте регистрации сочетаний клавиш для уже используемых надстроек, однако в некоторых случаях может потребоваться переопределить существующие сочетания клавиш или обработать конфликты между несколькими надстройкими, которые зарегистрировали одно и то же сочетание клавиш.

В случае конфликта пользователь будет видеть диалоговое окно при первой попытке использовать конфликтующей сочетания клавиш. Обратите внимание, что текст `name` для параметра надстройки, отображаемого в этом диалоговом окне, поступает из свойства в объекте действия в файле `shortcuts.json` .

![Иллюстрация, показывающая модальный конфликт с двумя разными действиями для одного ярлыка.](../images/add-in-shortcut-conflict-modal.png)

Пользователь может выбрать действие, которое будет выполнять сочетание клавиш. После выбора предпочтения сохраняются для дальнейшего использования того же ярлыка. Параметры ярлыка сохраняются для каждого пользователя на каждой платформе. Если пользователь хочет изменить параметры, он может вызвать команду **сброса** параметров ярлыка надстроек Office из поля поиска **"** Помощник". Вызов команды очищает все параметры ярлыка надстройки пользователя, и при следующей попытке использования конфликтующего ярлыка пользователю снова будет предложено ввести диалоговое окно конфликта.

![Поле поиска Помощника в Excel с действием сброса параметров ярлыка надстройки Office.](../images/add-in-reset-shortcuts-action.png)

Для оптимального взаимодействия с пользователем рекомендуется свести к минимуму конфликты с Excel с помощью этих рекомендаций.

- Используйте только сочетания клавиш со следующим шаблоном: **CTRL+SHIFT+ALT+* X***, где *x* является другим ключом.
- Если вам нужны дополнительные сочетания клавиш, проверьте список сочетаний клавиш [Excel](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f) и не используйте их в надстройке.
- Когда фокус клавиатуры находится в пользовательском интерфейсе надстройки, **CTRL+ПРОБЕЛ** и **CTRL+SHIFT+F10** не будут работать, так как это основные сочетания клавиш специальных возможностей.
- Если на компьютере с Windows или Mac команда "Сбросить параметры ярлыка надстроек Office" недоступна в меню поиска, пользователь может вручную добавить команду на ленту, настроив ленту через контекстное меню.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Настройка сочетаний клавиш для каждой платформы

Можно настроить ярлыки для конкретной платформы. Ниже приведен пример объекта, `shortcuts` который настраивает сочетания клавиш для каждой из следующих платформ: `windows`, `mac`, `web`. Обратите внимание, что у вас по-прежнему должен быть сочетание `default` клавиш для каждого ярлыка.

В следующем примере ключ является `default` резервным ключом для любой платформы, которая не указана. Не указана только платформа Windows, `default` поэтому ключ будет применяться только к Windows.

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

## <a name="localize-the-keyboard-shortcuts-json"></a>Локализация сочетаний клавиш JSON

Если надстройка поддерживает несколько языковых стандартов, `name` необходимо локализовать свойство объектов действия. Кроме того, если какой-либо из языковых стандартов, поддерживаемых надстройка, имеет разные алфавиты или системы записи и, следовательно, разные клавиатуры, вам также может потребоваться локализовать сочетания клавиш. Сведения о локализации сочетаний клавиш JSON см. в разделе ["Локализация расширенных переопределений"](../develop/localization.md#localize-extended-overrides).

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Ярлыки браузера, которые невозможно переопределить

При использовании настраиваемых сочетаний клавиш в Интернете некоторые сочетания клавиш, используемые браузером, не могут быть переопределены надстройки. Этот список выполняется. Если вы обнаруживаете другие сочетания, которые невозможно переопределить, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.

- CTRL+N
- CTRL+SHIFT+N
- CTRL+T
- CTRL+SHIFT+T
- CTRL+W
- Ctrl+PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users"></a>Включение настраиваемых сочетаний клавиш для определенных пользователей

Ваша надстройка позволяет пользователям переназначить действия надстройки альтернативным сочетаниям клавиатуры.

> [!NOTE]
> Для интерфейсов API, описанных в этом разделе, требуется набор обязательных элементов [KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) .

Используйте метод [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) , чтобы назначить пользовательские сочетания клавиатуры для действий надстроек. Метод принимает параметр `{[actionId:string]: string|null}`типа, `actionId`где s являются подмножеством идентификаторов действий, которые должны быть определены в JSON расширенного манифеста надстройки. Значения являются предпочтительным сочетанием ключей пользователя. Значение также может `null``actionId` быть , что приведет к удалению любой настройки для этого и возврату к сочетанию клавиатуры по умолчанию, которое определено в JSON расширенного манифеста надстройки.

Если пользователь вошел в Office, пользовательские сочетания сохраняются в перемещаемых параметрах пользователя для каждой платформы. Настройка ярлыков в настоящее время не поддерживается для анонимных пользователей.

```javascript
const userCustomShortcuts = {
    SHOWTASKPANE:"CTRL+SHIFT+1", 
    HIDETASKPANE:"CTRL+SHIFT+2"
};
Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(function () {
        console.log("Successfully registered.");
    })
    .catch(function (ex) {
        if (ex.code == "InvalidOperation") {
            console.log("ActionId does not exist or shortcut combination is invalid.");
        }
    });
```

Чтобы узнать, какие сочетания клавиш уже используются для пользователя, вызовите метод [Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member) . Этот метод возвращает объект типа `[actionId:string]:string|null}`, где значения представляют текущее сочетание клавиатуры, которое пользователь должен использовать для вызова указанного действия. Значения могут поступать из трех разных источников:

- Если возникл конфликт с ярлыком и пользователь выбрал другое действие (собственное или другое надстройка) для этого сочетания клавиатуры, `null` возвращаемое значение будет иметь значение, так как ярлык переопределен и пользователь в настоящее время не может использовать сочетание клавиатуры для вызова этого действия надстройки.
- Если ярлык был настроен с помощью метода [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) , возвращаемое значение будет настраиваемым сочетанием клавиатуры.
- Если ярлык не был переопределен или настроен, он возвращает значение из JSON расширенного манифеста надстройки.

Ниже приведен пример.

```javascript
Office.actions.getShortcuts()
    .then(function (userShortcuts) {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });

```

Как описано в [разделе "](#avoid-key-combinations-in-use-by-other-add-ins)Избегайте сочетаний клавиш, используемых другими надстройкими", рекомендуется избегать конфликтов в сочетаниях клавиш. Чтобы узнать, используется ли уже одно или несколько сочетаний клавиш, передайте их в виде массива строк в метод [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member) . Метод возвращает отчет, содержащий сочетания клавиш, которые уже используются в виде массива объектов типа `{shortcut: string, inUse: boolean}`. Свойство `shortcut` представляет собой сочетание клавиш, например CTRL+SHIFT+1. Если сочетание уже зарегистрировано в другом действии, свойству `inUse` задается значение `true`. Например, `[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]`. Ниже приведен пример фрагмента кода.

```javascript
const shortcuts = ["CTRL+SHIFT+1", "CTRL+SHIFT+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then(function (inUseArray) {
        const availableShortcuts = inUseArray.filter(function (shortcut) { return !shortcut.inUse; });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter(function (shortcut) { return shortcut.inUse; });
        console.log(usedShortcuts);
    });

```

## <a name="next-steps"></a>Дальнейшие действия

- См. пример [надстройки excel для](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) сочетаний клавиш.
- Общие сведения о работе с расширенными переопределениями в [Work с расширенными переопределениями манифеста](../develop/extended-overrides.md).
