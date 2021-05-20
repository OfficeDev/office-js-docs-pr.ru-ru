---
title: Создание пользовательских контекстуальных вкладок Office дополнительных надстройок
description: Узнайте, как добавить пользовательские контекстуальные вкладки в Office add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555208"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="4f058-103">Создание пользовательских контекстуальных вкладок Office дополнительных надстройок</span><span class="sxs-lookup"><span data-stu-id="4f058-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="4f058-104">Контекстная вкладка представляет 1000 скрытых элементов управления вкладками в ленте Office, которая отображается в строке вкладок при произном событии в Office документа.</span><span class="sxs-lookup"><span data-stu-id="4f058-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="4f058-105">Например, **вкладка «Дизайн** таблицы», которая отображается на Excel лентой при выборе таблицы.</span><span class="sxs-lookup"><span data-stu-id="4f058-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="4f058-106">Вы можете включить пользовательские контекстуальные вкладки в Office Add-in и указать, когда они видны или скрыты, создавая обработчики событий, которые меняют видимость.</span><span class="sxs-lookup"><span data-stu-id="4f058-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="4f058-107">(Однако пользовательские контекстуальные вкладки не реагируют на изменения фокусировки.)</span><span class="sxs-lookup"><span data-stu-id="4f058-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="4f058-108">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="4f058-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="4f058-109">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="4f058-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="4f058-110">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="4f058-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="4f058-111">Пользовательские контекстуальные вкладки в настоящее время поддерживаются только Excel и только на этих платформах и сборках:</span><span class="sxs-lookup"><span data-stu-id="4f058-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="4f058-112">Excel на Windows (Microsoft 365 подписка): Версия 2102 (Build 13801.20294) или позже.</span><span class="sxs-lookup"><span data-stu-id="4f058-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="4f058-113">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="4f058-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="4f058-114">Пользовательские контекстуальные вкладки работают только на платформах, хтякуя следующих наборов требований.</span><span class="sxs-lookup"><span data-stu-id="4f058-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="4f058-115">Для получения дополнительной информации о наборах требований и о том, как с ними [работать, Office приложений и требований API.](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="4f058-116">РиббонАпи 1.2</span><span class="sxs-lookup"><span data-stu-id="4f058-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="4f058-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="4f058-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="4f058-118">Вы можете использовать проверки времени выполнения в коде, чтобы проверить, поддерживает ли комбинация хоста и платформы пользователя эти наборы требований, [описанные в приложениях Office и API.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="4f058-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="4f058-119">(Техника указания наборов требований в манифесте, которая также описана в этой статье, в настоящее время не работает для RibbonApi 1.2.) Кроме того, можно реализовать [альтернативный пользовательский интерфейс, когда пользовательские контекстуальные вкладки не поддерживаются.](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)</span><span class="sxs-lookup"><span data-stu-id="4f058-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="4f058-120">Поведение пользовательских контекстуальных вкладок</span><span class="sxs-lookup"><span data-stu-id="4f058-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="4f058-121">Пользовательский интерфейс для пользовательских контекстуальных вкладок следует шаблону встроенных Office контекстуальных вкладок.</span><span class="sxs-lookup"><span data-stu-id="4f058-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="4f058-122">Ниже приведены основные принципы для размещения пользовательских контекстуальных вкладок:</span><span class="sxs-lookup"><span data-stu-id="4f058-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="4f058-123">Когда видна пользовательская контекстуальная вкладка, она отображается на правом конце ленты.</span><span class="sxs-lookup"><span data-stu-id="4f058-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="4f058-124">Если одновременно видны одна или несколько встроенных контекстуальных вкладок и одна или несколько пользовательских контекстуальных вкладок из надстройок, пользовательские контекстуальные вкладки всегда справа от всех встроенных контекстуальных вкладок.</span><span class="sxs-lookup"><span data-stu-id="4f058-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="4f058-125">Если надстройка имеет несколько контекстуальных вкладок и есть контексты, в которых видно несколько, они отображаются в порядке, в котором они определены в надстройке.</span><span class="sxs-lookup"><span data-stu-id="4f058-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="4f058-126">(Направление в том же направлении, что и Office язык; то есть слева направо на языках слева направо, но справа налево на языках справа налево.) Определите [группы и элементы управления, которые появляются на вкладке, для получения](#define-the-groups-and-controls-that-appear-on-the-tab) подробной информации о том, как вы их определяете.</span><span class="sxs-lookup"><span data-stu-id="4f058-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="4f058-127">Если более одного дополнения имеет контекстуальную вкладку, видимую в определенном контексте, то они отображаются в порядке запуска надстройок.</span><span class="sxs-lookup"><span data-stu-id="4f058-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="4f058-128">Пользовательские *контекстуальные* вкладки, в отличие от пользовательских основных вкладок, не добавляются Office к ленте приложения.</span><span class="sxs-lookup"><span data-stu-id="4f058-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="4f058-129">Они присутствуют только в Office документах, на которых работает надстройку.</span><span class="sxs-lookup"><span data-stu-id="4f058-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="4f058-130">Основные шаги для включения контекстной вкладки в надстройку</span><span class="sxs-lookup"><span data-stu-id="4f058-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="4f058-131">Ниже приведены основные шаги для включения пользовательской контекстуальной вкладки в дополнение:</span><span class="sxs-lookup"><span data-stu-id="4f058-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="4f058-132">Настройте надстройу для использования общего времени выполнения.</span><span class="sxs-lookup"><span data-stu-id="4f058-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="4f058-133">Определите вкладку и группы и элементы управления, которые появляются на ней.</span><span class="sxs-lookup"><span data-stu-id="4f058-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="4f058-134">Зарегистрируйте контекстную вкладку с помощью Office.</span><span class="sxs-lookup"><span data-stu-id="4f058-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="4f058-135">Укажите обстоятельства, при которых вкладка будет видна.</span><span class="sxs-lookup"><span data-stu-id="4f058-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="4f058-136">Настройка надстройок для использования общего времени выполнения</span><span class="sxs-lookup"><span data-stu-id="4f058-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="4f058-137">Добавление пользовательских контекстуальных вкладок требует, чтобы надстройка использовалась в общем времени выполнения.</span><span class="sxs-lookup"><span data-stu-id="4f058-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="4f058-138">Для получения дополнительной [информации см.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="4f058-139">Определите группы и элементы управления, которые отображаются на вкладке</span><span class="sxs-lookup"><span data-stu-id="4f058-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="4f058-140">В отличие от пользовательских основных вкладок, которые определяются с XML в манифесте, пользовательские контекстуальные вкладки определяются во время выполнения с каплей JSON.</span><span class="sxs-lookup"><span data-stu-id="4f058-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="4f058-141">Код анализирует каплю в объект JavaScript, а затем передает объект в [метод Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="4f058-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="4f058-142">Пользовательские контекстуальные вкладки присутствуют только в документах, на которых в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="4f058-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="4f058-143">Это отличается от пользовательских основных вкладок, которые добавляются в ленту приложения Office, когда надстройка установлена и остается присутствовать при открытом другом документе.</span><span class="sxs-lookup"><span data-stu-id="4f058-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="4f058-144">Кроме того, `requestCreateControls` метод может быть запущен только один раз в сеансе вашего дополнения.</span><span class="sxs-lookup"><span data-stu-id="4f058-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="4f058-145">Если он вызван снова, ошибка брошена.</span><span class="sxs-lookup"><span data-stu-id="4f058-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="4f058-146">Структура свойств и подпредложений капли JSON (и ключевых имен) примерно параллельна структуре [элемента CustomTab и его](../reference/manifest/customtab.md) потомкам в манифесте XML.</span><span class="sxs-lookup"><span data-stu-id="4f058-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="4f058-147">Мы построим пример контекстуальных вкладок JSON blob шаг за шагом.</span><span class="sxs-lookup"><span data-stu-id="4f058-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="4f058-148">Полная схема контекстной вкладки JSON находится [dynamic-ribbon.schema.jsна](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="4f058-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="4f058-149">Если вы работаете в Visual Studio Code, вы можете использовать этот файл, чтобы IntelliSense и проверить ваш JSON.</span><span class="sxs-lookup"><span data-stu-id="4f058-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="4f058-150">Для получения дополнительной информации [см Visual Studio Code.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)</span><span class="sxs-lookup"><span data-stu-id="4f058-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="4f058-151">Начните с создания строки JSON с двумя свойствами массива, названными `actions` и `tabs` .</span><span class="sxs-lookup"><span data-stu-id="4f058-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="4f058-152">Массив `actions` является спецификацией всех функций, которые могут быть выполнены с помощью элементов управления контекстуальной вкладкой. Массив `tabs` определяет одну или несколько контекстуальных *вкладок, максимум до 20.*</span><span class="sxs-lookup"><span data-stu-id="4f058-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="4f058-153">Этот простой пример контекстуальной вкладки будет иметь только одну кнопку и, таким образом, только одно действие.</span><span class="sxs-lookup"><span data-stu-id="4f058-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="4f058-154">Добавьте следующее в качестве единственного члена `actions` массива.</span><span class="sxs-lookup"><span data-stu-id="4f058-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="4f058-155">Об этой разметке, обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f058-155">About this markup, note:</span></span>

    - <span data-ttu-id="4f058-156">Свойства `id` `type` и свойства являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="4f058-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="4f058-157">Значение может быть `type` либо "ExecuteFunction" или "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="4f058-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="4f058-158">Свойство `functionName` используется только тогда, когда значение `type` `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="4f058-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="4f058-159">Это название функции, определенной в FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="4f058-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="4f058-160">Для получения дополнительной информации о FunctionFile [см. Основные концепции для дополнительных команд.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="4f058-161">На более позднем этапе вы сопоставите это действие с кнопкой на контекстуальной вкладке.</span><span class="sxs-lookup"><span data-stu-id="4f058-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="4f058-162">Добавьте следующее в качестве единственного члена `tabs` массива.</span><span class="sxs-lookup"><span data-stu-id="4f058-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="4f058-163">Об этой разметке, обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f058-163">About this markup, note:</span></span>

    - <span data-ttu-id="4f058-164">Свойство `id` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="4f058-164">The `id` property is required.</span></span> <span data-ttu-id="4f058-165">Используйте краткий, описательный идентификатор, который уникален среди всех контекстуальных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="4f058-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="4f058-166">Свойство `label` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="4f058-166">The `label` property is required.</span></span> <span data-ttu-id="4f058-167">Это удобный строка, чтобы служить в качестве метки контекстуальной вкладке.</span><span class="sxs-lookup"><span data-stu-id="4f058-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="4f058-168">Свойство `groups` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="4f058-168">The `groups` property is required.</span></span> <span data-ttu-id="4f058-169">Он определяет группы элементов управления, которые будут отображаться на вкладке. Он должен иметь по крайней мере *одного члена и не более 20*.</span><span class="sxs-lookup"><span data-stu-id="4f058-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="4f058-170">(Есть также ограничения на количество элементов управления, которые можно иметь на пользовательской контекстной вкладке, и что также будет ограничивать, сколько групп, которые у вас есть.</span><span class="sxs-lookup"><span data-stu-id="4f058-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="4f058-171">Дополнительную информацию можно посмотреть на следующий шаг.)</span><span class="sxs-lookup"><span data-stu-id="4f058-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="4f058-172">Объект вкладки также может иметь дополнительное `visible` свойство, которое определяет, видна ли вкладка сразу же при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="4f058-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="4f058-173">Поскольку контекстуальные вкладки обычно скрыты до тех пор, пока событие пользователя не запустит видимость (например, пользователь выбирает сущность того или иного типа в документе), `visible` свойство по умолчанию, `false` когда его нет.</span><span class="sxs-lookup"><span data-stu-id="4f058-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="4f058-174">В более позднем разделе мы покажем, как настроить свойство `true` в ответ на событие.</span><span class="sxs-lookup"><span data-stu-id="4f058-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="4f058-175">В простом постоянном примере контекстуальная вкладка имеет только одну группу.</span><span class="sxs-lookup"><span data-stu-id="4f058-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="4f058-176">Добавьте следующее в качестве единственного члена `groups` массива.</span><span class="sxs-lookup"><span data-stu-id="4f058-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="4f058-177">Об этой разметке, обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f058-177">About this markup, note:</span></span>

    - <span data-ttu-id="4f058-178">Все свойства необходимы.</span><span class="sxs-lookup"><span data-stu-id="4f058-178">All the properties are required.</span></span>
    - <span data-ttu-id="4f058-179">Свойство должно быть уникальным среди всех групп во `id` вкладке. Используйте краткое, описательное удостоверение личности.</span><span class="sxs-lookup"><span data-stu-id="4f058-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="4f058-180">Это `label` удобный строка, чтобы служить в качестве ярлыка группы.</span><span class="sxs-lookup"><span data-stu-id="4f058-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="4f058-181">Значение `icon` свойства – это массив объектов, которые определяют значки, которые группа будет иметь на ленте в зависимости от размера ленты и Office окна приложения.</span><span class="sxs-lookup"><span data-stu-id="4f058-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="4f058-182">Значение `controls` свойства – это массив объектов, которые определяют кнопки и меню в группе.</span><span class="sxs-lookup"><span data-stu-id="4f058-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="4f058-183">Должен быть хотя бы один.</span><span class="sxs-lookup"><span data-stu-id="4f058-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="4f058-184">*Общее количество элементов управления на всей вкладке может быть не более 20.*</span><span class="sxs-lookup"><span data-stu-id="4f058-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="4f058-185">Например, можно иметь 3 группы по 6 элементов управления и четвертую группу с 2 элементами управления, но вы не можете иметь 4 группы по 6 элементов управления каждая.</span><span class="sxs-lookup"><span data-stu-id="4f058-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="4f058-186">Каждая группа должна иметь значок не менее двух размеров, 32x32 px и 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="4f058-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="4f058-187">Дополнительно, вы также можете иметь значки размеров 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, и 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="4f058-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="4f058-188">Office, какой значок использовать в зависимости от размера ленты и Office окна приложения.</span><span class="sxs-lookup"><span data-stu-id="4f058-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="4f058-189">Добавьте следующие объекты в массив значков.</span><span class="sxs-lookup"><span data-stu-id="4f058-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="4f058-190">(Если размеры окна и ленты достаточно велики, чтобы по крайней мере один из *элементов* управления в группе не появлялся, то значок группы вообще не отображается.</span><span class="sxs-lookup"><span data-stu-id="4f058-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="4f058-191">Например, наблюдайте за **группой** стилей на ленте Word при сокращении и расширении окна Word.) Об этой разметке, обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f058-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="4f058-192">Оба свойства необходимы.</span><span class="sxs-lookup"><span data-stu-id="4f058-192">Both the properties are required.</span></span>
    - <span data-ttu-id="4f058-193">Единица `size` измерения свойства пикселей.</span><span class="sxs-lookup"><span data-stu-id="4f058-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="4f058-194">Иконки всегда квадратные, поэтому число и высота, и ширина.</span><span class="sxs-lookup"><span data-stu-id="4f058-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="4f058-195">Свойство `sourceLocation` указывает полный URL на значок.</span><span class="sxs-lookup"><span data-stu-id="4f058-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="4f058-196">Точно так же, как обычно необходимо изменить URL-адреса в манифесте дополнения при переходе от разработки к производству (например, изменение домена с локального хостинга на contoso.com), вы также должны изменить URL-адреса в контекстуальных вкладок JSON.</span><span class="sxs-lookup"><span data-stu-id="4f058-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="4f058-197">В нашем простом постоянном примере у группы есть только одна кнопка.</span><span class="sxs-lookup"><span data-stu-id="4f058-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="4f058-198">Добавьте следующий объект в качестве единственного члена `controls` массива.</span><span class="sxs-lookup"><span data-stu-id="4f058-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="4f058-199">Об этой разметке, обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="4f058-199">About this markup, note:</span></span>

    - <span data-ttu-id="4f058-200">Все свойства, кроме , `enabled` не требуется.</span><span class="sxs-lookup"><span data-stu-id="4f058-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="4f058-201">`type` определяет тип управления.</span><span class="sxs-lookup"><span data-stu-id="4f058-201">`type` specifies the type of control.</span></span> <span data-ttu-id="4f058-202">Значения могут быть "Кнопка", "Меню", или "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="4f058-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="4f058-203">`id` может быть до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="4f058-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="4f058-204">`actionId` должен быть идентификатор действия, определенный в `actions` массиве.</span><span class="sxs-lookup"><span data-stu-id="4f058-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="4f058-205">(См. шаг 1 этого раздела.)</span><span class="sxs-lookup"><span data-stu-id="4f058-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="4f058-206">`label` является удобной строкой, которая служит в качестве метки кнопки.</span><span class="sxs-lookup"><span data-stu-id="4f058-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="4f058-207">`superTip` представляет собой богатую форму наконечника инструмента.</span><span class="sxs-lookup"><span data-stu-id="4f058-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="4f058-208">Требуются `title` как `description` свойства, так и свойства.</span><span class="sxs-lookup"><span data-stu-id="4f058-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="4f058-209">`icon` указывает значки для кнопки.</span><span class="sxs-lookup"><span data-stu-id="4f058-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="4f058-210">Предыдущие замечания о значке группы применимы и здесь.</span><span class="sxs-lookup"><span data-stu-id="4f058-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="4f058-211">`enabled` (необязательно) определяет, включена ли кнопка при запуске контекстуальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="4f058-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="4f058-212">По умолчанию, если нет `true` .</span><span class="sxs-lookup"><span data-stu-id="4f058-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="4f058-213">Ниже приводится полный пример капли JSON:</span><span class="sxs-lookup"><span data-stu-id="4f058-213">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="4f058-214">Зарегистрируйте контекстную вкладку с помощью Office с запросомCreateControls</span><span class="sxs-lookup"><span data-stu-id="4f058-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="4f058-215">Контекстная вкладка регистрируется с Office, [позвонив по Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="4f058-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="4f058-216">Обычно это делается либо в функции, назначенной `Office.initialize` методу, либо с `Office.onReady` ним.</span><span class="sxs-lookup"><span data-stu-id="4f058-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="4f058-217">Для получения дополнительной информации об этих методах и инициализации надстройок см [Office.](../develop/initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="4f058-218">Однако можно позвонить в метод в любое время после инициализации.</span><span class="sxs-lookup"><span data-stu-id="4f058-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4f058-219">Метод `requestCreateControls` может быть вызван только один раз в данном сеансе дополнения.</span><span class="sxs-lookup"><span data-stu-id="4f058-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="4f058-220">Ошибка брошена, если она вызвана снова.</span><span class="sxs-lookup"><span data-stu-id="4f058-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="4f058-221">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4f058-221">The following is an example.</span></span> <span data-ttu-id="4f058-222">Обратите внимание, что строка JSON должна быть преобразована в объект JavaScript с `JSON.parse` помощью метода, прежде чем она может быть передана функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4f058-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="4f058-223">Укажите контексты, когда вкладка будет видна с запросомUpdate</span><span class="sxs-lookup"><span data-stu-id="4f058-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="4f058-224">Как правило, пользовательская контекстуальная вкладка должна отображаться, когда событие, инициированное пользователем, изменяет контекст дополнения.</span><span class="sxs-lookup"><span data-stu-id="4f058-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="4f058-225">Рассмотрим сценарий, в котором вкладка должна быть видна, когда и только когда активируется диаграмма (на листе Excel рабочей книги по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="4f058-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="4f058-226">Начните с назначения обработчиков.</span><span class="sxs-lookup"><span data-stu-id="4f058-226">Begin by assigning handlers.</span></span> <span data-ttu-id="4f058-227">Обычно это делается в `Office.onReady` методе, как в следующем примере, который присваивает обработчикам (созданным на более позднем `onActivated` этапе) и событиям всех диаграмм `onDeactivated` в листе.</span><span class="sxs-lookup"><span data-stu-id="4f058-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="4f058-228">Далее определите обработчиков.</span><span class="sxs-lookup"><span data-stu-id="4f058-228">Next, define the handlers.</span></span> <span data-ttu-id="4f058-229">Ниже приводится простой пример `showDataTab` , но см [Обработка HostRestartNeeded ошибка позже](#handle-the-hostrestartneeded-error) в этой статье для более надежной версии функции.</span><span class="sxs-lookup"><span data-stu-id="4f058-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="4f058-230">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="4f058-230">About this code, note:</span></span>

- <span data-ttu-id="4f058-231">Office определяет время обновления состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="4f058-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="4f058-232">Метод [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) выстроил в очередь запрос на обновление.</span><span class="sxs-lookup"><span data-stu-id="4f058-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="4f058-233">Метод разрешит объект, `Promise` как только он выстроился в очередь с запросом, а не когда лента фактически обновляется.</span><span class="sxs-lookup"><span data-stu-id="4f058-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="4f058-234">Параметром метода `requestUpdate` является объект [RibbonUpdaterData,](/javascript/api/office/office.ribbonupdaterdata) который (1) определяет вкладку по своему *идентификатору точно так, как* указано в JSON и (2) определяет видимость вкладки.</span><span class="sxs-lookup"><span data-stu-id="4f058-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="4f058-235">Если у вас есть несколько пользовательских контекстуальных вкладок, которые должны быть видны в том же контексте, вы просто добавить дополнительные объекты вкладок в `tabs` массив.</span><span class="sxs-lookup"><span data-stu-id="4f058-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="4f058-236">Обработчик, чтобы скрыть вкладку почти идентичны, за исключением того, что он `visible` устанавливает свойство обратно `false` .</span><span class="sxs-lookup"><span data-stu-id="4f058-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="4f058-237">Библиотека Office JavaScript также предоставляет несколько интерфейсов (типов), чтобы упростить строительство `RibbonUpdateData` объекта.</span><span class="sxs-lookup"><span data-stu-id="4f058-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="4f058-238">Ниже приводится `showDataTab` функция в TypeScript и он использует эти типы.</span><span class="sxs-lookup"><span data-stu-id="4f058-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="4f058-239">Переключение видимости вкладок и включенного состояния кнопки одновременно</span><span class="sxs-lookup"><span data-stu-id="4f058-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="4f058-240">Метод `requestUpdate` также используется для переключения включенного или отключенного статуса пользовательской кнопки на пользовательской контекстуальной вкладке или пользовательской вкладке ядра. Для получения подробной информации об этом [см.](disable-add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="4f058-241">Могут быть сценарии, в которых вы хотите изменить как видимость вкладки, так и состояние кнопки одновременно.</span><span class="sxs-lookup"><span data-stu-id="4f058-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="4f058-242">Вы можете сделать это с помощью одного звонка `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="4f058-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="4f058-243">Ниже приводится пример, в котором кнопка на вкладке ядра включена в то же время, как контекстуальная вкладка сделана видимой.</span><span class="sxs-lookup"><span data-stu-id="4f058-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="4f058-244">В следующем примере кнопка, включенная, находится на той же контекстной вкладке, которая делается видимой.</span><span class="sxs-lookup"><span data-stu-id="4f058-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="4f058-245">Локализация капли JSON</span><span class="sxs-lookup"><span data-stu-id="4f058-245">Localizing the JSON blob</span></span>

<span data-ttu-id="4f058-246">Bb JSON, передаваемый, не локализован так `requestCreateControls` же, как локализовка манифеста для пользовательских основных вкладок (которая [описана при локализации управления из манифеста).](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="4f058-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="4f058-247">Вместо этого локализация должна происходить во время выполнения с использованием различных капли JSON для каждого места.</span><span class="sxs-lookup"><span data-stu-id="4f058-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="4f058-248">Мы предлагаем использовать `switch` выписку, которая тестирует [Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="4f058-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="4f058-249">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4f058-249">The following is an example:</span></span>

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

<span data-ttu-id="4f058-250">Затем код вызывает функцию, чтобы получить локализованную каплю, которая `requestCreateControls` передается, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="4f058-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="4f058-251">Лучшие практики для пользовательских контекстуальных вкладок</span><span class="sxs-lookup"><span data-stu-id="4f058-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="4f058-252">Реализация альтернативного пользовательского интерфейса при поддержке пользовательских контекстуальных вкладок</span><span class="sxs-lookup"><span data-stu-id="4f058-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="4f058-253">Некоторые комбинации платформы, Office приложения и Office сборки не `requestCreateControls` поддерживают.</span><span class="sxs-lookup"><span data-stu-id="4f058-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="4f058-254">Надстройа должна быть разработана таким образом, чтобы предоставить альтернативный опыт пользователям, которые работают надстройок на одной из этих комбинаций.</span><span class="sxs-lookup"><span data-stu-id="4f058-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="4f058-255">В следующих разделах описаны два способа обеспечения возврата.</span><span class="sxs-lookup"><span data-stu-id="4f058-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="4f058-256">Использование неконтекстуальных вкладок или элементов управления</span><span class="sxs-lookup"><span data-stu-id="4f058-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="4f058-257">Существует явный элемент, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), который предназначен для создания обратного опыта в дополнение, которое реализует пользовательские контекстуальные вкладки, когда надстройка работает на приложении или платформе, которая не поддерживает пользовательские контекстуальные вкладки.</span><span class="sxs-lookup"><span data-stu-id="4f058-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="4f058-258">Простейшей стратегией использования этого элемента является определение в манифесте одной или нескольких пользовательских основных вкладок (то есть *неконтекстуальных* пользовательских вкладок), которые дублируют настройки ленты пользовательских контекстуальных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="4f058-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="4f058-259">Но вы `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` добавляете в качестве первого элемента ребенка [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="4f058-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="4f058-260">Результатом этого является следующее:</span><span class="sxs-lookup"><span data-stu-id="4f058-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="4f058-261">Если надстройка выполняется на приложении и платформе, ею поддерживают пользовательские контекстуальные вкладки, то пользовательская вкладка ядра не будет отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="4f058-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="4f058-262">Вместо этого пользовательская контекстуальная вкладка будет создана при вызове `requestCreateControls` надстройки метода.</span><span class="sxs-lookup"><span data-stu-id="4f058-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="4f058-263">Если надстройка выполняется на приложении или платформе, *которая не поддерживает,* `requestCreateControls` то пользовательская вкладка ядра появляется на ленте.</span><span class="sxs-lookup"><span data-stu-id="4f058-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="4f058-264">Ниже приводится пример этой простой стратегии.</span><span class="sxs-lookup"><span data-stu-id="4f058-264">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="4f058-265">Эта простая стратегия использует пользовательскую основную вкладку, которая отражает пользовательскую контекстуальную вкладку с ее детскими группами и управлениями, но вы можете использовать более сложную стратегию.</span><span class="sxs-lookup"><span data-stu-id="4f058-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="4f058-266">Элемент `<OverriddenByRibbonApi>` также может быть добавлен в качестве (первого) элемента ребенка в элементы [группы и](../reference/manifest/group.md) [управления](../reference/manifest/control.md) (как тип кнопки, [так и](../reference/manifest/control.md#button-control) тип [меню),](../reference/manifest/control.md#menu-dropdown-button-controls)а также элементы `<Item>` меню.</span><span class="sxs-lookup"><span data-stu-id="4f058-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="4f058-267">Этот факт позволяет распространять группы и элементы управления, которые в противном случае появились бы на контекстной вкладке между различными группами, кнопками и меню в различных пользовательских основных вкладок.</span><span class="sxs-lookup"><span data-stu-id="4f058-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="4f058-268">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="4f058-268">The following is an example.</span></span> <span data-ttu-id="4f058-269">Обратите внимание, что "MyButton" появится на пользовательской вкладке ядра только тогда, когда пользовательские контекстуальные вкладки не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="4f058-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="4f058-270">Но родительская группа и пользовательская вкладка ядра будут отображаться независимо от того, поддерживаются ли пользовательские контекстуальные вкладки.</span><span class="sxs-lookup"><span data-stu-id="4f058-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="4f058-271">Дополнительные примеры [см.](../reference/manifest/overriddenbyribbonapi.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="4f058-272">Когда родительские вкладки, группы или `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` меню отмечены, то это не видно, и все это ребенка разметки игнорируется, когда пользовательские контекстуальные вкладки не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="4f058-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="4f058-273">Таким образом, не имеет значения, если какой-либо из этих элементов ребенка `<OverriddenByRibbonApi>` элемент или то, что его значение.</span><span class="sxs-lookup"><span data-stu-id="4f058-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="4f058-274">Смысл этого заключается в том, что если элемент меню, элемент управления или группа должны быть видны во всех контекстах, то не только он не должен быть отмечен `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` , но *его родоначальницы меню, группы и вкладки также не должны быть отмечены таким образом*.</span><span class="sxs-lookup"><span data-stu-id="4f058-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4f058-275">Не *помекайте* все элементы ребенка вкладкой, группой или `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` меню.</span><span class="sxs-lookup"><span data-stu-id="4f058-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="4f058-276">Это бессмысленно, если родительский элемент помечен по `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` причинам, уявечаемым в предыдущем пункте.</span><span class="sxs-lookup"><span data-stu-id="4f058-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="4f058-277">Кроме того, если вы `<OverriddenByRibbonApi>` оставите на родителей (или установить `false` его), то родитель появится независимо от того, пользовательские контекстуальные вкладки поддерживаются, но он будет пуст, когда они поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="4f058-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="4f058-278">Таким образом, если все элементы ребенка не должны отображаться при поддержке пользовательских контекстуальных вкладок, отметь родительский и только `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` родительский.</span><span class="sxs-lookup"><span data-stu-id="4f058-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="4f058-279">Используйте API, которые показывают или скрывают панели задач в определенных контекстах</span><span class="sxs-lookup"><span data-stu-id="4f058-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="4f058-280">В качестве `<OverriddenByRibbonApi>` альтернативы, ваше дополнение может определить панели задач с пользовательским интерфейсом элементов управления, которые дублируют функциональность элементов управления на пользовательских контекстуальных вкладку. Затем используйте [методы Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) [и Office.addin.hide,](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) чтобы показать панели задач, когда и только когда контекстуальная вкладка была бы показана, если бы она была поддержана.</span><span class="sxs-lookup"><span data-stu-id="4f058-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="4f058-281">Для получения подробной информации о том, как использовать [эти методы, см. Показать или скрыть панели задач Office add-in.](../develop/show-hide-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="4f058-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="4f058-282">Ручка HostRestartNeededed ошибка</span><span class="sxs-lookup"><span data-stu-id="4f058-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="4f058-283">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="4f058-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="4f058-284">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="4f058-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="4f058-285">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="4f058-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="4f058-286">Ваш код должен обрабатывать эту ошибку.</span><span class="sxs-lookup"><span data-stu-id="4f058-286">Your code should handle this error.</span></span> <span data-ttu-id="4f058-287">Ниже приводится пример того, как.</span><span class="sxs-lookup"><span data-stu-id="4f058-287">The following is an example of how.</span></span> <span data-ttu-id="4f058-288">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="4f058-288">In this case, the `reportError` method displays the error to the user.</span></span>

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
