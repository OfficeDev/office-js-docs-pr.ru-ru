---
title: Пользовательские ярлыки клавиатуры в надстройки Office
description: Узнайте, как добавить в надстройку Office настраиваемые клавиши, также известные как сочетания клавиш.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505201"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="68617-103">Добавление ярлыков настраиваемой клавиатуры в надстройки Office (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="68617-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="68617-104">Ярлыки клавиатуры, также известные как сочетания клавиш, позволяют пользователям надстройки работать эффективнее и они улучшают доступность надстройки для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.</span><span class="sxs-lookup"><span data-stu-id="68617-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="68617-105">Чтобы начать с рабочей версии надстройки с уже включенными клавишами, клонировать и запускать примеры ярлыков [клавиатуры Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="68617-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="68617-106">Если вы готовы добавить ярлыки клавиатуры в собственную надстройку, продолжи эту статью.</span><span class="sxs-lookup"><span data-stu-id="68617-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="68617-107">Существует три шага, чтобы добавить в надстройку ярлыки клавиатуры:</span><span class="sxs-lookup"><span data-stu-id="68617-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="68617-108">[Настройка манифеста надстройки.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="68617-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="68617-109">[Создание или изменение ярлыков JSON-файла для](#create-or-edit-the-shortcuts-json-file) определения действий и их клавиш.</span><span class="sxs-lookup"><span data-stu-id="68617-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="68617-110">[Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить функцию с каждым действием.</span><span class="sxs-lookup"><span data-stu-id="68617-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="68617-111">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="68617-111">Configure the manifest</span></span>

<span data-ttu-id="68617-112">В манифест необходимо внести два небольших изменения.</span><span class="sxs-lookup"><span data-stu-id="68617-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="68617-113">Один из них — включить надстройку для использования общего времени работы, а другой — указать на файл в формате JSON, в котором определены ярлыки клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68617-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="68617-114">Настройка надстройки для использования общего времени работы</span><span class="sxs-lookup"><span data-stu-id="68617-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="68617-115">Добавление пользовательских ярлыков клавиатуры требует от надстройки использовать общее время работы.</span><span class="sxs-lookup"><span data-stu-id="68617-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="68617-116">Дополнительные сведения: [Настройка надстройки для использования общего времени работы.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="68617-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="68617-117">Привязка файла сопоставления к манифесту</span><span class="sxs-lookup"><span data-stu-id="68617-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="68617-118">Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="68617-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="68617-119">Установите атрибут для полного URL-адреса файла JSON в проекте, который будет создан `Url` на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="68617-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="68617-120">Создание или изменение ярлыков JSON-файла</span><span class="sxs-lookup"><span data-stu-id="68617-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="68617-121">Создайте файл JSON в проекте.</span><span class="sxs-lookup"><span data-stu-id="68617-121">Create a JSON file in your project.</span></span> <span data-ttu-id="68617-122">Убедитесь, что путь файла соответствует расположению, указанному для атрибута элемента `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="68617-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="68617-123">В этом файле будут описаны ярлыки клавиатуры и действия, которые они будут вызывать.</span><span class="sxs-lookup"><span data-stu-id="68617-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="68617-124">В файле JSON есть два массива.</span><span class="sxs-lookup"><span data-stu-id="68617-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="68617-125">Массив действий будет содержать объекты, которые определяют действия, которые будут вызываться, а массив ярлыков будет содержать объекты, которые соотносят комбинации ключей с действиями.</span><span class="sxs-lookup"><span data-stu-id="68617-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="68617-126">Вот пример:</span><span class="sxs-lookup"><span data-stu-id="68617-126">Here is an example:</span></span>

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
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="68617-127">Дополнительные сведения об объектах JSON см. в дополнительных сведениях [о](#constructing-the-action-objects) том, как создавать объекты действия и создавать [объекты ярлыка.](#constructing-the-shortcut-objects)</span><span class="sxs-lookup"><span data-stu-id="68617-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="68617-128">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68617-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="68617-129">В этой статье можно использовать "CONTROL" на месте "CTRL".</span><span class="sxs-lookup"><span data-stu-id="68617-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="68617-130">На более позднем этапе действия сами будут соедему с функциями, которые вы пишете.</span><span class="sxs-lookup"><span data-stu-id="68617-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="68617-131">В этом примере вы позже назовет SHOWTASKPANE функцией, которая вызывает метод, а HIDETASKPANE — функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.</span><span class="sxs-lookup"><span data-stu-id="68617-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="68617-132">Создание сопоставления действий с их функциями</span><span class="sxs-lookup"><span data-stu-id="68617-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="68617-133">В проекте откройте файл JavaScript, загруженный вашей htmL-страницей в `<FunctionFile>` элементе.</span><span class="sxs-lookup"><span data-stu-id="68617-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="68617-134">В файле JavaScript используйте [API Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы составить карту каждого действия, указанного в файле JSON, с функцией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68617-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="68617-135">Добавьте в файл следующий JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68617-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="68617-136">Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="68617-136">Note the following about the code:</span></span>

    - <span data-ttu-id="68617-137">Первый параметр — это одно из действий из файла JSON.</span><span class="sxs-lookup"><span data-stu-id="68617-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="68617-138">Второй параметр — это функция, которая выполняется при нажатии клавиши на сочетание ключей, относясь к действию в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="68617-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="68617-139">Чтобы продолжить пример, используйте `'SHOWTASKPANE'` в качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="68617-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="68617-140">Чтобы открыть область задач надстройки, используйте метод [Office.addin.showTaskpane.](/javascript/api/office/office.addin#showastaskpane--)</span><span class="sxs-lookup"><span data-stu-id="68617-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="68617-141">После этого код должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="68617-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="68617-142">Добавьте второй вызов `Office.actions.associate` функции, чтобы соединить действие с функцией, вызываемой `HIDETASKPANE` [Office.addin.hide.](/javascript/api/office/office.addin#hide--)</span><span class="sxs-lookup"><span data-stu-id="68617-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="68617-143">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68617-143">The following is an example:</span></span>

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

<span data-ttu-id="68617-144">После предыдущих действий надстройка позволяет переключать видимость области задач, нажав клавишу **стрелки Ctrl+Shift+Up** и клавишу **стрелки Ctrl+Shift+Down.**</span><span class="sxs-lookup"><span data-stu-id="68617-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="68617-145">Это такое же поведение, как показано в примере надстройки [клавиш Excel клавиатуры](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="68617-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="68617-146">Сведения и ограничения</span><span class="sxs-lookup"><span data-stu-id="68617-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="68617-147">Построение объектов действия</span><span class="sxs-lookup"><span data-stu-id="68617-147">Constructing the action objects</span></span>

<span data-ttu-id="68617-148">Используйте следующие рекомендации при указании объектов в массиве `action` shortcuts.js:</span><span class="sxs-lookup"><span data-stu-id="68617-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="68617-149">Имена свойств `id` и `name` обязательны.</span><span class="sxs-lookup"><span data-stu-id="68617-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="68617-150">Свойство `id` используется для уникальной идентификации действия, вызываемого с помощью ярлыка клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68617-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="68617-151">Свойство `name` должно быть удобной строкой, описываемой действием.</span><span class="sxs-lookup"><span data-stu-id="68617-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="68617-152">Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="68617-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="68617-153">Свойство `type`— необязательное.</span><span class="sxs-lookup"><span data-stu-id="68617-153">The `type` property is optional.</span></span> <span data-ttu-id="68617-154">В `ExecuteFunction` настоящее время поддерживается только тип.</span><span class="sxs-lookup"><span data-stu-id="68617-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="68617-155">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68617-155">The following is an example:</span></span>

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

<span data-ttu-id="68617-156">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68617-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="68617-157">Построение объектов ярлыка</span><span class="sxs-lookup"><span data-stu-id="68617-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="68617-158">Используйте следующие рекомендации при указании объектов в массиве `shortcuts` shortcuts.js:</span><span class="sxs-lookup"><span data-stu-id="68617-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="68617-159">Имена свойств `action` `key` и `default` обязательно.</span><span class="sxs-lookup"><span data-stu-id="68617-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="68617-160">Значение свойства является строкой и должно соответствовать одному из свойств `action` `id` объекта действия.</span><span class="sxs-lookup"><span data-stu-id="68617-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="68617-161">Свойство может быть любым сочетанием символов A - Z, a -z, 0 - 9, а знаки препинания `default` "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="68617-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="68617-162">(По соглашению в этих свойствах не используются буквы более низкого уровня.)</span><span class="sxs-lookup"><span data-stu-id="68617-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="68617-163">Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (ALT, CTRL, SHIFT) и только одного ключа.</span><span class="sxs-lookup"><span data-stu-id="68617-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="68617-164">Для macs мы также поддерживаем ключ модификатора COMMAND.</span><span class="sxs-lookup"><span data-stu-id="68617-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="68617-165">Для Компьютеров Mac ALT соедем на клавишу OPTION.</span><span class="sxs-lookup"><span data-stu-id="68617-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="68617-166">Для Windows командная команда относит к клавише CTRL.</span><span class="sxs-lookup"><span data-stu-id="68617-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="68617-167">Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, ALT+a и ALT+A являются одним и тем же ярлыком, как `default` и CTRL+- и CTRL+, так как "-" и "_" являются одним и тем же физическим \_ ключом.</span><span class="sxs-lookup"><span data-stu-id="68617-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="68617-168">Символ "+" указывает, что клавиши с обеих сторон нажаты одновременно.</span><span class="sxs-lookup"><span data-stu-id="68617-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="68617-169">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="68617-169">The following is an example:</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

<span data-ttu-id="68617-170">Полная схема для ярлыков JSON находится [вextended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="68617-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="68617-171">Клавиши, также известные как последовательное клавиши, например ярлык Excel для выбора цвета заполнения **Alt+H, H,** не поддерживаются в надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="68617-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="68617-172">Использование ярлыков, когда фокус находится в области задач</span><span class="sxs-lookup"><span data-stu-id="68617-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="68617-173">В настоящее время ярлыки клавиатуры для надстройки Office можно вызывать только в том случае, если фокус пользователя находится в таблице.</span><span class="sxs-lookup"><span data-stu-id="68617-173">Currently, the keyboard shortcuts for an Office Add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="68617-174">Если фокус пользователя находится внутри пользовательского интерфейса Office (например, области задач), ни один из ярлыков надстройки не игнорируется.</span><span class="sxs-lookup"><span data-stu-id="68617-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="68617-175">В качестве обхода надстройка может определять обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится внутри пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="68617-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="68617-176">Использование комбинаций ключей, которые уже используются Office или другой надстройки</span><span class="sxs-lookup"><span data-stu-id="68617-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="68617-177">Во время предварительного просмотра не существует системы определения того, что происходит, когда пользователь нажимает клавишу, зарегистрированную надстройка, а также Office или другой надстройки.</span><span class="sxs-lookup"><span data-stu-id="68617-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="68617-178">Поведение неопределяется.</span><span class="sxs-lookup"><span data-stu-id="68617-178">Behavior is undefined.</span></span>

<span data-ttu-id="68617-179">В настоящее время не существует обхода, когда два или несколько надстроек зарегистрировали один и тот же ярлык клавиатуры, но можно минимизировать конфликты с Excel с помощью этих методов:</span><span class="sxs-lookup"><span data-stu-id="68617-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="68617-180">Используйте только клавиши со следующим шаблоном в надстройки: \**Ctrl+Shift+Alt+* x\*\*\*, где *x* — это другой ключ.</span><span class="sxs-lookup"><span data-stu-id="68617-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="68617-181">Если вам нужно больше клавиш, проверьте список ярлыков [клавиатуры Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и не использовать их в надстройки.</span><span class="sxs-lookup"><span data-stu-id="68617-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="68617-182">Ярлыки браузера, которые нельзя переопределять</span><span class="sxs-lookup"><span data-stu-id="68617-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="68617-183">Вы не можете использовать ни одной из следующих комбинаций клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="68617-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="68617-184">Они используются браузерами и не могут быть переопределены.</span><span class="sxs-lookup"><span data-stu-id="68617-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="68617-185">Этот список находится в процессе выполнения.</span><span class="sxs-lookup"><span data-stu-id="68617-185">This list is a work in progress.</span></span> <span data-ttu-id="68617-186">Если вы обнаружите другие комбинации, которые нельзя переопределять, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.</span><span class="sxs-lookup"><span data-stu-id="68617-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="68617-187">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="68617-187">Ctrl+N</span></span>
- <span data-ttu-id="68617-188">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="68617-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="68617-189">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="68617-189">Ctrl+T</span></span>
- <span data-ttu-id="68617-190">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="68617-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="68617-191">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="68617-191">Ctrl+W</span></span>
- <span data-ttu-id="68617-192">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="68617-192">Ctrl+PgUp/PgDn</span></span>

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="68617-193">Локализовать ярлыки клавиатуры JSON</span><span class="sxs-lookup"><span data-stu-id="68617-193">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="68617-194">Если надстройка поддерживает несколько локалов, необходимо локализовать свойство `name` объектов действия.</span><span class="sxs-lookup"><span data-stu-id="68617-194">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="68617-195">Кроме того, если в любом из локаутах, поддерживаюх надстройку, есть алфавиты или различные системы записи, а значит, и другие клавиатуры, возможно, потребуется также локализовать ярлыки.</span><span class="sxs-lookup"><span data-stu-id="68617-195">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="68617-196">Сведения о том, как локализовать клавиши ярлыков JSON, см. в рубрезе [Localize extended overrides.](../develop/localization.md#localize-extended-overrides)</span><span class="sxs-lookup"><span data-stu-id="68617-196">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="next-steps"></a><span data-ttu-id="68617-197">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="68617-197">Next Steps</span></span>

- <span data-ttu-id="68617-198">См. пример [надстройки excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="68617-198">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
- <span data-ttu-id="68617-199">Получите обзор работы с расширенными переопределениями в Работе с расширенными [переопределениями манифеста.](../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="68617-199">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
