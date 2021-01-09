---
title: Настраиваемые сочетания клавиш в надстройки Office
description: Узнайте, как добавить в надстройку Office пользовательские сочетания клавиш, также известные как сочетания клавиш.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: dc99674b92ebb415b1d49fb28821d8c2e34c8077
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789151"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="f8ae4-103">Добавление пользовательских сочетания клавиш в надстройки Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="f8ae4-104">Сочетания клавиш, также известные как сочетания клавиш, позволяют пользователям надстройки работать эффективнее и улучшают ее доступность для пользователей с ограниченными возможностями, предоставляя альтернативу мыши.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="f8ae4-105">Чтобы начать с рабочей версии надстройки с уже включенными сочетаниями клавиш, клонировать и запустить пример сочетания клавиш [Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="f8ae4-106">Когда вы будете готовы добавить сочетания клавиш в вашу собственную надстройку, продолжайте работу с этой статьей.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="f8ae4-107">Чтобы добавить сочетания клавиш в надстройку, необходимо три шага:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="f8ae4-108">[Настройте манифест надстройки.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="f8ae4-109">[Создайте или отредактируетЕ JSON-файл](#create-or-edit-the-shortcuts-json-file) ярлыков, чтобы определить действия и их сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="f8ae4-110">[Добавьте один или несколько вызовов](#create-a-mapping-of-actions-to-their-functions) API [Office.actions.associate](/javascript/api/office/office.actions#associate) во время работы, чтобы соотоставить функцию с каждым действием.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f8ae4-111">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="f8ae4-111">Configure the manifest</span></span>

<span data-ttu-id="f8ae4-112">Манифесту необходимо внести два небольших изменения.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="f8ae4-113">Первый — позволить надстройки использовать общую времени работы, а другой — указать файл в формате JSON, в котором вы определили сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="f8ae4-114">Настройка надстройки для использования общей времени работы</span><span class="sxs-lookup"><span data-stu-id="f8ae4-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="f8ae4-115">Для добавления настраиваемого сочетания клавиш надстройка использует общую времени работы.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="f8ae4-116">Для получения дополнительных [сведений настройте надстройку](../develop/configure-your-add-in-to-use-a-shared-runtime.md)для использования общей времени работы.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="f8ae4-117">Привязка файла сопоставления к манифесту</span><span class="sxs-lookup"><span data-stu-id="f8ae4-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="f8ae4-118">Непосредственно *под* элементом манифеста (не внутри) добавьте `<VersionOverrides>` элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="f8ae4-119">Установите для атрибута полный URL-адрес JSON-файла в проекте, который будет создан на `Url` более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="f8ae4-120">Создание или изменение JSON-файла ярлыков</span><span class="sxs-lookup"><span data-stu-id="f8ae4-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="f8ae4-121">Создайте JSON-файл в проекте.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-121">Create a JSON file in your project.</span></span> <span data-ttu-id="f8ae4-122">Убедитесь, что путь к файлу соответствует расположению, указанному для `Url` атрибута [элемента ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="f8ae4-123">В этом файле описываются сочетания клавиш и действия, которые они будут вызывать.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="f8ae4-124">Внутри JSON-файла есть два массива.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="f8ae4-125">Массив действий будет содержать объекты, которые определяют действия, которые необходимо вызвать, а массив ярлыков будет содержать объекты, которые соотнося сочетания клавиш с действиями.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="f8ae4-126">Вот пример:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-126">Here is an example:</span></span>

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

    <span data-ttu-id="f8ae4-127">Дополнительные сведения об объектах JSON см. в подстройки объектов [действий](#constructing-the-action-objects) и создания объектов [ярлыков.](#constructing-the-shortcut-objects)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="f8ae4-128">Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="f8ae4-129">В этой статье вы можете использовать control, а не CTRL.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="f8ae4-130">На более позднем этапе эти действия будут сами сописываться с функциями, которые вы пишете.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="f8ae4-131">В этом примере вы позже созовем SHOWTASKPANE с функцией, которая вызывает метод, и HIDETASKPANE с функцией, которая `Office.addin.showAsTaskpane` вызывает `Office.addin.hide` метод.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="f8ae4-132">Создание сопоставления действий с их функциями</span><span class="sxs-lookup"><span data-stu-id="f8ae4-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="f8ae4-133">В проекте откройте файл JavaScript, загруженный HTML-страницей в `<FunctionFile>` элементе.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="f8ae4-134">В файле JavaScript используйте API [Office.actions.associate,](/javascript/api/office/office.actions#associate) чтобы соотоставить каждое действие, указанное в файле JSON, с функцией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="f8ae4-135">Добавьте в файл следующий javaScript.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="f8ae4-136">Обратите внимание на следующие вопросы о коде:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-136">Note the following about the code:</span></span>

    - <span data-ttu-id="f8ae4-137">Первый параметр — это одно из действий из JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="f8ae4-138">Второй параметр — это функция, которая запускается, когда пользователь нажмет сочетание клавиш, которое со карты с действием в JSON-файле.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="f8ae4-139">Чтобы продолжить пример, используйте `'SHOWTASKPANE'` его в качестве первого параметра.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="f8ae4-140">Для тела функции используйте метод [Office.addin.showTaskpane,](/javascript/api/office/office.addin#showastaskpane--) чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="f8ae4-141">После этого код должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="f8ae4-142">Добавьте второй вызов `Office.actions.associate` функции, чтобы соединить действие `HIDETASKPANE` с функцией, которая вызывает [Office.addin.hide.](/javascript/api/office/office.addin#hide--)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="f8ae4-143">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-143">The following is an example:</span></span>

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

<span data-ttu-id="f8ae4-144">После предыдущих действий надстройка может переключать видимость области задач, нажимая **клавиши CTRL+SHIFT+СТРЕЛКА** ВВЕРХ и **CTRL+SHIFT+СТРЕЛКА ВНИЗ.**</span><span class="sxs-lookup"><span data-stu-id="f8ae4-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="f8ae4-145">Это поведение такое же, как показано в примере надстройки [Excel с сочетаниями клавиш.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="f8ae4-146">Сведения и ограничения</span><span class="sxs-lookup"><span data-stu-id="f8ae4-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="f8ae4-147">Создание объектов действий</span><span class="sxs-lookup"><span data-stu-id="f8ae4-147">Constructing the action objects</span></span>

<span data-ttu-id="f8ae4-148">При указании объектов в массиве shortcuts.jsиспользуйте `action` следующие рекомендации:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="f8ae4-149">Имена свойств `id` являются `name` обязательными.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="f8ae4-150">Свойство используется для уникальной идентификации `id` действия, вызываемого с помощью сочетания клавиш.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="f8ae4-151">Свойство `name` должно быть пользовательской строкой, описывающий действие.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="f8ae4-152">Это должно быть сочетание символов A - Z, a - z, 0 - 9 и знаков препинания "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="f8ae4-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="f8ae4-153">Свойство `type`— необязательное.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-153">The `type` property is optional.</span></span> <span data-ttu-id="f8ae4-154">В `ExecuteFunction` настоящее время поддерживается только тип.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="f8ae4-155">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-155">The following is an example:</span></span>

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

<span data-ttu-id="f8ae4-156">Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="f8ae4-157">Создание объектов ярлыков</span><span class="sxs-lookup"><span data-stu-id="f8ae4-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="f8ae4-158">При указании объектов в массиве shortcuts.jsиспользуйте `shortcuts` следующие рекомендации:</span><span class="sxs-lookup"><span data-stu-id="f8ae4-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="f8ae4-159">Имена свойств `action` , и являются `key` `default` обязательной.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="f8ae4-160">Значение свойства является строкой и должно соответствовать одному из `action` `id` свойств в объекте действия.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="f8ae4-161">Свойство может быть любым сочетанием символов `default` A - Z, a -z, 0 -9 и знаками препинания "-", "_" и "+".</span><span class="sxs-lookup"><span data-stu-id="f8ae4-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="f8ae4-162">(По соглашению буквы нижнего дела не используются в этих свойствах.)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="f8ae4-163">Свойство должно содержать имя по крайней мере одного ключа модификатора `default` (ALT, CTRL, SHIFT) и только одного другого ключа.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="f8ae4-164">Для Mac также поддерживается клавиша модификатора COMMAND.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="f8ae4-165">Для Компьютеров Mac ALT со картой OPTION.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="f8ae4-166">Для Windows command со карты с клавишей CTRL.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="f8ae4-167">Если два символа связаны с одним и тем же физическим ключом в стандартной клавиатуре, они являются синонимами в свойстве; например, ALT+a и ALT+A — это один и тот же ярлык, как `default` и CTRL+- и CTRL+, так как "-" и "_" являются одним и тем же физическим \_ ключом.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="f8ae4-168">Символ "+" указывает, что клавиши с обеих сторон нажимаются одновременно.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="f8ae4-169">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-169">The following is an example:</span></span>

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

<span data-ttu-id="f8ae4-170">Полная схема для ярлыков JSON находится в [extended-manifest.schema.js.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="f8ae4-171">Клавиши, также известные как последовательное сочетания клавиш, например ярлык Excel для выбора цвета заливки **ALT+H, H,** не поддерживаются в надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="f8ae4-172">Использование ярлыков, когда фокус находится в области задач</span><span class="sxs-lookup"><span data-stu-id="f8ae4-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="f8ae4-173">В настоящее время сочетания клавиш для надстройки Office можно вызывать только в том случае, если фокус пользователя находится на этом планшете.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-173">Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="f8ae4-174">Если фокус пользователя находится в пользовательском интерфейсе Office (например, в области задач), ни один из ярлыков надстройки не игнорируется.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="f8ae4-175">В качестве обходного решения надстройка может определять обработчики клавиатуры, которые могут вызывать определенные действия, когда фокус пользователя находится в пользовательском интерфейсе надстройки.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="f8ae4-176">Использование сочетаний клавиш, которые уже используются Office или другой надстройки</span><span class="sxs-lookup"><span data-stu-id="f8ae4-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="f8ae4-177">Во время предварительного просмотра не существует системы для определения того, что происходит, когда пользователь нажимает сочетание клавиш, которое зарегистрировано надстройки, а также Office или другой надстройки.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="f8ae4-178">Поведение неопределяется.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-178">Behavior is undefined.</span></span>

<span data-ttu-id="f8ae4-179">В настоящее время обходной путь не существует, если две или более надстроек зарегистрировали один и тот же ярлык клавиатуры, но вы можете свести к минимуму конфликты с Excel, выполив указанные здесь действия.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="f8ae4-180">Используйте в надстройки только сочетания клавиш со следующим шаблоном: \**CTRL+SHIFT+ALT+* x\*\*\*, где *x* — это другой ключ.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="f8ae4-181">Если вам нужны дополнительные сочетания клавиш, проверьте список сочетания клавиш [Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)и избегайте их использования в надстройки.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="f8ae4-182">Ярлыки браузера, которые не могут быть переопределены</span><span class="sxs-lookup"><span data-stu-id="f8ae4-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="f8ae4-183">Нельзя использовать любое из следующих сочетаний клавиатуры.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="f8ae4-184">Они используются браузерами и не могут быть переопределены.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="f8ae4-185">Этот список является ходом выполнения.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-185">This list is a work in progress.</span></span> <span data-ttu-id="f8ae4-186">Если вы обнаружите другие сочетания, которые не могут быть переопределены, сообщите нам об этом с помощью средства обратной связи в нижней части этой страницы.</span><span class="sxs-lookup"><span data-stu-id="f8ae4-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="f8ae4-187">CTRL+N</span><span class="sxs-lookup"><span data-stu-id="f8ae4-187">Ctrl+N</span></span>
- <span data-ttu-id="f8ae4-188">CTRL+SHIFT+N</span><span class="sxs-lookup"><span data-stu-id="f8ae4-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="f8ae4-189">CTRL+T</span><span class="sxs-lookup"><span data-stu-id="f8ae4-189">Ctrl+T</span></span>
- <span data-ttu-id="f8ae4-190">CTRL+SHIFT+T</span><span class="sxs-lookup"><span data-stu-id="f8ae4-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="f8ae4-191">CTRL+W</span><span class="sxs-lookup"><span data-stu-id="f8ae4-191">Ctrl+W</span></span>
- <span data-ttu-id="f8ae4-192">CTRL+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="f8ae4-192">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="f8ae4-193">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="f8ae4-193">Next Steps</span></span>

- <span data-ttu-id="f8ae4-194">См. пример надстройки [Excel-keyboard-shortcuts.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="f8ae4-194">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
