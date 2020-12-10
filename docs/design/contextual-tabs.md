---
title: Создание настраиваемых контекстных вкладок в надстройках Office
description: Узнайте, как добавить настраиваемые контекстные вкладки в надстройку Office.
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: d8617c7dd8748d15393c0e38c527062e5894e791
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612738"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="895a8-103">Создание настраиваемых контекстных вкладок в надстройках Office (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="895a8-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="895a8-104">Контекстная вкладка — это скрытый элемент управления "Вкладка" на ленте Office, который отображается в строке вкладок при возникновении указанного события в документе Office.</span><span class="sxs-lookup"><span data-stu-id="895a8-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="895a8-105">Например, вкладка **конструктор таблицы** , которая отображается на ленте Excel при выборе таблицы.</span><span class="sxs-lookup"><span data-stu-id="895a8-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="895a8-106">Вы можете включить настраиваемые контекстные вкладки в надстройку Office и указать, когда они видимы или скрыты, создавая обработчики событий, которые изменяют видимость.</span><span class="sxs-lookup"><span data-stu-id="895a8-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="895a8-107">(Однако настраиваемые контекстные вкладки не отвечают на изменения фокуса.)</span><span class="sxs-lookup"><span data-stu-id="895a8-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="895a8-108">В этой статье предполагается, что вы уже ознакомились с приведенной ниже документацией.</span><span class="sxs-lookup"><span data-stu-id="895a8-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="895a8-109">Просмотрите ее, если вы работали с командами надстроек (настраиваемыми элементами меню и кнопками ленты) некоторое время назад.</span><span class="sxs-lookup"><span data-stu-id="895a8-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="895a8-110">Основные концепции команд надстроек</span><span class="sxs-lookup"><span data-stu-id="895a8-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="895a8-111">Настраиваемые контекстные вкладки доступны в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="895a8-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="895a8-112">Поэкспериментируйте с ними в среде разработки или тестирования, но не добавляйте их в производственную надстройку.</span><span class="sxs-lookup"><span data-stu-id="895a8-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="895a8-113">В настоящее время Настраиваемые контекстные вкладки поддерживаются только в Excel и только на этих платформах и построениях:</span><span class="sxs-lookup"><span data-stu-id="895a8-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="895a8-114">Excel в Windows (только для Microsoft 365, а не Бессрочная лицензия): версия 2011 (сборка 13426,20274).</span><span class="sxs-lookup"><span data-stu-id="895a8-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="895a8-115">Возможно, ваша подписка на Microsoft 365 должна быть включена в [текущий канал (Предварительная версия)](https://insider.office.com/join/windows) , ранее называемый "Monthly Channel (targeted)" или "Предварительная оценка".</span><span class="sxs-lookup"><span data-stu-id="895a8-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="895a8-116">Настраиваемые контекстные вкладки работают только на платформах, поддерживающих следующие наборы требований.</span><span class="sxs-lookup"><span data-stu-id="895a8-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="895a8-117">Дополнительные сведения о наборах требований и способах работы с ними можно узнать в статье [Указание приложений Office и требований к API](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="895a8-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="895a8-118">Шаредрунтиме 1,1</span><span class="sxs-lookup"><span data-stu-id="895a8-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="895a8-119">Поведение настраиваемых контекстных вкладок</span><span class="sxs-lookup"><span data-stu-id="895a8-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="895a8-120">Пользовательский интерфейс для настраиваемых контекстных вкладок соответствует шаблону встроенных контекстных вкладок Office.</span><span class="sxs-lookup"><span data-stu-id="895a8-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="895a8-121">Ниже приведены основные принципы для настраиваемых контекстных вкладок размещения.</span><span class="sxs-lookup"><span data-stu-id="895a8-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="895a8-122">Если пользовательская контекстная вкладка отображается, она отображается в правой части ленты.</span><span class="sxs-lookup"><span data-stu-id="895a8-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="895a8-123">Если одновременно отображаются одна или несколько встроенных контекстных вкладок и одна или несколько настраиваемых контекстных вкладок из надстроек, пользовательские контекстные вкладки всегда расположены справа от всех встроенных контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="895a8-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="895a8-124">Если у надстройки есть несколько контекстных вкладок и есть контексты, в которых отображаются несколько контекстов, они отображаются в том порядке, в котором они определены в надстройке.</span><span class="sxs-lookup"><span data-stu-id="895a8-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="895a8-125">(Направление — это то же направление, что и язык Office; то есть слева направо для языков с письмом слева направо, но справа налево для языков с письмом справа налево.) [В разделе Определение групп и элементов управления, которые отображаются на вкладке,](#define-the-groups-and-controls-that-appear-on-the-tab) для получения сведений о том, как они определены.</span><span class="sxs-lookup"><span data-stu-id="895a8-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="895a8-126">Если в нескольких надстройках есть контекстная вкладка, видимая в определенном контексте, они отображаются в том порядке, в котором были запущены надстройки.</span><span class="sxs-lookup"><span data-stu-id="895a8-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="895a8-127">Настраиваемые *Контекстные* вкладки, в отличие от пользовательских основных вкладок, не добавляются в ленту приложения Office без возможности восстановления.</span><span class="sxs-lookup"><span data-stu-id="895a8-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="895a8-128">Они существуют только в документах Office, в которых работает ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="895a8-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="895a8-129">Основные действия по добавлению контекстной вкладки в надстройке</span><span class="sxs-lookup"><span data-stu-id="895a8-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="895a8-130">Ниже приведены основные действия по добавлению настраиваемой контекстной вкладки в надстройке.</span><span class="sxs-lookup"><span data-stu-id="895a8-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="895a8-131">Настройка надстройки для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="895a8-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="895a8-132">Определите вкладки, а также группы и элементы управления, которые отображаются на ней.</span><span class="sxs-lookup"><span data-stu-id="895a8-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="895a8-133">Зарегистрируйте контекстную вкладку в Office.</span><span class="sxs-lookup"><span data-stu-id="895a8-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="895a8-134">Укажите обстоятельства, когда вкладка станет видимой.</span><span class="sxs-lookup"><span data-stu-id="895a8-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="895a8-135">Настройка надстройки для использования общей среды выполнения</span><span class="sxs-lookup"><span data-stu-id="895a8-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="895a8-136">Для добавления настраиваемых контекстных вкладок необходимо, чтобы ваша надстройка использовала общую среду выполнения.</span><span class="sxs-lookup"><span data-stu-id="895a8-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="895a8-137">Дополнительные сведения см. в статье [Настройка надстройки для использования общей среды выполнения](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="895a8-137">For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="895a8-138">Определение групп и элементов управления, которые отображаются на вкладке</span><span class="sxs-lookup"><span data-stu-id="895a8-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="895a8-139">В отличие от пользовательских основных вкладок, которые определены с помощью XML в манифесте, Настраиваемые контекстные вкладки определяются во время выполнения с помощью большого двоичного объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="895a8-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="895a8-140">Ваш код анализирует большой двоичный объект в объект JavaScript, а затем передает объект в метод [Office. Ribbon. рекуесткреатеконтролс](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) .</span><span class="sxs-lookup"><span data-stu-id="895a8-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="895a8-141">Настраиваемые контекстные вкладки присутствуют только в документах, в которых в данный момент запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="895a8-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="895a8-142">Это отличается от настраиваемых основных вкладок, которые добавляются на ленту приложений Office при установке и оставлении презентации при открытии другого документа.</span><span class="sxs-lookup"><span data-stu-id="895a8-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="895a8-143">Кроме того, `requestCreateControls` метод можно выполнить только один раз в сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="895a8-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="895a8-144">Если он вызывается повторно, возникает ошибка.</span><span class="sxs-lookup"><span data-stu-id="895a8-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="895a8-145">Структура свойств и вложенных объектов объекта BLOB в JSON (и имена ключей) приблизительно параллельна структуре элемента [CustomTab](../reference/manifest/customtab.md) и его потомков в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="895a8-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="895a8-146">Мы создадим пример пошагового двоичного объекта JSON для контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="895a8-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="895a8-147">(Полная схема для контекстной вкладки JSON [dynamic-ribbon.schema.jsвключена](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="895a8-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="895a8-148">Эта ссылка может не работать в течение раннего периода предварительной версии для контекстных вкладок.</span><span class="sxs-lookup"><span data-stu-id="895a8-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="895a8-149">Если ссылка не работает, вы можете найти последний черновик схемы в [dynamic-ribbon.schema.jsчерновиков](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Если вы работаете в Visual Studio Code, вы можете использовать этот файл для получения IntelliSense и проверки JSON.</span><span class="sxs-lookup"><span data-stu-id="895a8-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="895a8-150">Дополнительные сведения см в статье [Редактирование JSON с помощью схем и параметров Visual Studio Code: JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="895a8-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="895a8-151">Сначала создайте строку JSON с двумя свойствами массива Names `actions` и `tabs` .</span><span class="sxs-lookup"><span data-stu-id="895a8-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="895a8-152">`actions`Массив — это спецификация всех функций, которые могут быть выполнены элементами управления на контекстной вкладке. `tabs`Массив определяет одну или несколько контекстных вкладок, *максимум до 10*.</span><span class="sxs-lookup"><span data-stu-id="895a8-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="895a8-153">Этот простой пример контекстной вкладки будет содержать только одну кнопку и, таким образом, только одно действие.</span><span class="sxs-lookup"><span data-stu-id="895a8-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="895a8-154">Добавьте следующий элемент в качестве единственного элемента `actions` массива.</span><span class="sxs-lookup"><span data-stu-id="895a8-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="895a8-155">Об этой разметке Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="895a8-155">About this markup, note:</span></span>

    - <span data-ttu-id="895a8-156">`id`Свойства и `type` являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="895a8-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="895a8-157">`type`Возможные значения: "ExecuteFunction" или "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="895a8-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="895a8-158">`functionName`Свойство используется только в том случае, если значение `type` равно `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="895a8-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="895a8-159">Это имя функции, определенной в элементе FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="895a8-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="895a8-160">Более подробную информацию о FunctionFile можно узнать в статье [Основные понятия для команд надстроек](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="895a8-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="895a8-161">На следующем этапе вы добавите это действие к кнопке на контекстной вкладке.</span><span class="sxs-lookup"><span data-stu-id="895a8-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="895a8-162">Добавьте следующий элемент в качестве единственного элемента `tabs` массива.</span><span class="sxs-lookup"><span data-stu-id="895a8-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="895a8-163">Об этой разметке Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="895a8-163">About this markup, note:</span></span>

    - <span data-ttu-id="895a8-164">Свойство `id` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="895a8-164">The `id` property is required.</span></span> <span data-ttu-id="895a8-165">Используйте краткий описательный идентификатор, который является уникальным среди всех контекстных вкладок в надстройке.</span><span class="sxs-lookup"><span data-stu-id="895a8-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="895a8-166">Свойство `label` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="895a8-166">The `label` property is required.</span></span> <span data-ttu-id="895a8-167">Это понятная для пользователя строка, которая выступает в качестве метки контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="895a8-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="895a8-168">Свойство `groups` является обязательным.</span><span class="sxs-lookup"><span data-stu-id="895a8-168">The `groups` property is required.</span></span> <span data-ttu-id="895a8-169">Он определяет группы элементов управления, которые будут отображаться на вкладке. У нее должен быть по крайней мере один участник *и не более 20*.</span><span class="sxs-lookup"><span data-stu-id="895a8-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="895a8-170">(Существуют также ограничения на количество элементов управления, которые можно использовать на настраиваемой контекстной вкладке, и это также ограничит количество групп, которые у вас есть.</span><span class="sxs-lookup"><span data-stu-id="895a8-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="895a8-171">Для получения дополнительных сведений просмотрите следующий шаг.</span><span class="sxs-lookup"><span data-stu-id="895a8-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="895a8-172">Объект Tab также может иметь необязательное `visible` свойство, которое указывает, будет ли вкладка отображаться немедленно при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="895a8-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="895a8-173">Так как контекстные вкладки обычно скрываются до тех пор, пока пользователь не назначит их видимость (например, пользователь выбирает сущность некоторого типа в документе), `visible` свойство по умолчанию имеет значение, `false` если оно не задано.</span><span class="sxs-lookup"><span data-stu-id="895a8-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="895a8-174">В более позднем разделе мы покажем, как установить свойство `true` в ответ на событие.</span><span class="sxs-lookup"><span data-stu-id="895a8-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="895a8-175">В простом примере контекстная вкладка содержит только одну группу.</span><span class="sxs-lookup"><span data-stu-id="895a8-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="895a8-176">Добавьте следующий элемент в качестве единственного элемента `groups` массива.</span><span class="sxs-lookup"><span data-stu-id="895a8-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="895a8-177">Об этой разметке Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="895a8-177">About this markup, note:</span></span>

    - <span data-ttu-id="895a8-178">Все свойства являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="895a8-178">All the properties are required.</span></span>
    - <span data-ttu-id="895a8-179">`id`Свойство должно быть уникальным среди всех групп на вкладке. Используйте краткий описательный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="895a8-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="895a8-180">`label`Представляет собой удобную для пользователя строку, которая выступает в качестве метки группы.</span><span class="sxs-lookup"><span data-stu-id="895a8-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="895a8-181">`icon`Значение свойства — это массив объектов, указывающий значки, которые будут находиться на ленте в зависимости от размера ленты и окна приложения Office.</span><span class="sxs-lookup"><span data-stu-id="895a8-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="895a8-182">`controls`Значение свойства это массив объектов, указывающих кнопки и меню в группе.</span><span class="sxs-lookup"><span data-stu-id="895a8-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="895a8-183">В группе должно быть по крайней мере один и *не более 6*.</span><span class="sxs-lookup"><span data-stu-id="895a8-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="895a8-184">*Общее количество элементов управления на вкладке целиком не может превышать 20.*</span><span class="sxs-lookup"><span data-stu-id="895a8-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="895a8-185">Например, у вас есть 3 группы с 6 элементами управления и четвертая группа с 2 элементами управления, но у вас не может быть 4 группы с 6 элементами управления.</span><span class="sxs-lookup"><span data-stu-id="895a8-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="895a8-186">У каждой группы должен быть по крайней мере два размера: 32x32 px и 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="895a8-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="895a8-187">Кроме того, можно использовать значки размеров 16x16 точек, 20x20 px, 24x24 px, 40x40 px, 48x48 px и 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="895a8-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="895a8-188">Office определяет, какой значок следует использовать в зависимости от размера ленты и окна приложения Office.</span><span class="sxs-lookup"><span data-stu-id="895a8-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="895a8-189">Добавьте следующие объекты в массив значков.</span><span class="sxs-lookup"><span data-stu-id="895a8-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="895a8-190">(Если размер окна и ленты достаточно велик для отображения по крайней мере одного из *элементов управления* в группе, значок группы не отображается.</span><span class="sxs-lookup"><span data-stu-id="895a8-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="895a8-191">В качестве примера просмотрите группу **styles** на ленте Word, когда вы сжимаете и разворачиваете окно Word. Об этой разметке Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="895a8-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="895a8-192">Оба свойства являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="895a8-192">Both the properties are required.</span></span>
    - <span data-ttu-id="895a8-193">`size`Единицей измерения свойства является точка.</span><span class="sxs-lookup"><span data-stu-id="895a8-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="895a8-194">Значки всегда квадратны, поэтому число будет равно высоте и ширине.</span><span class="sxs-lookup"><span data-stu-id="895a8-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="895a8-195">`sourceLocation`Свойство указывает полный URL-адрес значка.</span><span class="sxs-lookup"><span data-stu-id="895a8-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="895a8-196">Как правило, необходимо изменить URL-адреса в манифесте надстройки при переходе от разработки к рабочей среде (например, при изменении домена с localhost на contoso.com), необходимо также изменить URL-адреса в контекстных вкладках JSON.</span><span class="sxs-lookup"><span data-stu-id="895a8-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="895a8-197">В нашем простом примере группа содержит только одну кнопку.</span><span class="sxs-lookup"><span data-stu-id="895a8-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="895a8-198">Добавьте указанный ниже объект в качестве единственного элемента `controls` массива.</span><span class="sxs-lookup"><span data-stu-id="895a8-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="895a8-199">Об этой разметке Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="895a8-199">About this markup, note:</span></span>

    - <span data-ttu-id="895a8-200">Все свойства, кроме `enabled` , обязательны.</span><span class="sxs-lookup"><span data-stu-id="895a8-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="895a8-201">`type` Указывает тип элемента управления.</span><span class="sxs-lookup"><span data-stu-id="895a8-201">`type` specifies the type of control.</span></span> <span data-ttu-id="895a8-202">Возможные значения: "Button", "Menu" или "Мобилебуттон".</span><span class="sxs-lookup"><span data-stu-id="895a8-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="895a8-203">`id` может содержать до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="895a8-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="895a8-204">`actionId` должен быть ИДЕНТИФИКАТОРом действия, определенным в `actions` массиве.</span><span class="sxs-lookup"><span data-stu-id="895a8-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="895a8-205">(См. шаг 1 этого раздела).</span><span class="sxs-lookup"><span data-stu-id="895a8-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="895a8-206">`label` — Это понятная для пользователя строка, которая выступает в качестве метки кнопки.</span><span class="sxs-lookup"><span data-stu-id="895a8-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="895a8-207">`superTip` представляет полнофункциональную форму подсказки.</span><span class="sxs-lookup"><span data-stu-id="895a8-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="895a8-208">`title`Требуются и свойства, и `description` .</span><span class="sxs-lookup"><span data-stu-id="895a8-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="895a8-209">`icon` задает значки для кнопки.</span><span class="sxs-lookup"><span data-stu-id="895a8-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="895a8-210">Приведенные выше примечания относительно значка группы также применимы.</span><span class="sxs-lookup"><span data-stu-id="895a8-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="895a8-211">`enabled` (необязательно) указывает, включена ли кнопка при появлении контекстной вкладки.</span><span class="sxs-lookup"><span data-stu-id="895a8-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="895a8-212">Если значение не указано, используется значение по умолчанию `true` .</span><span class="sxs-lookup"><span data-stu-id="895a8-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="895a8-213">Ниже приведен полный пример большого двоичного объекта JSON:</span><span class="sxs-lookup"><span data-stu-id="895a8-213">The following is the complete example of the JSON blob:</span></span>

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
      "label": "Data",
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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="895a8-214">Регистрация контекстной вкладки с Office с помощью Рекуесткреатеконтролс</span><span class="sxs-lookup"><span data-stu-id="895a8-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="895a8-215">Контекстная вкладка регистрируется в Office, вызывая метод [Office. Ribbon. рекуесткреатеконтролс](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) .</span><span class="sxs-lookup"><span data-stu-id="895a8-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="895a8-216">Обычно это выполняется в функции, которая назначена `Office.initialize` или `Office.onReady` методу.</span><span class="sxs-lookup"><span data-stu-id="895a8-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="895a8-217">Дополнительные сведения об этих методах и инициализации надстройки приведены в статье [Initialize The Office Your Your](../develop/initialize-add-in.md)надстройка.</span><span class="sxs-lookup"><span data-stu-id="895a8-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="895a8-218">Тем не менее, вы можете вызвать метод в любое время после инициализации.</span><span class="sxs-lookup"><span data-stu-id="895a8-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="895a8-219">`requestCreateControls`Метод можно вызвать только один раз в данном сеансе надстройки.</span><span class="sxs-lookup"><span data-stu-id="895a8-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="895a8-220">При повторном вызове возникает ошибка.</span><span class="sxs-lookup"><span data-stu-id="895a8-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="895a8-221">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="895a8-221">The following is an example.</span></span> <span data-ttu-id="895a8-222">Обратите внимание, что строку JSON необходимо преобразовать в объект JavaScript с `JSON.parse` методом, прежде чем его можно будет передать в функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="895a8-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="895a8-223">Указание контекстов, когда вкладка будет отображаться в Рекуеступдате</span><span class="sxs-lookup"><span data-stu-id="895a8-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="895a8-224">Как правило, пользовательская контекстная вкладка отображается при изменении контекста надстройки событием, инициированным пользователем.</span><span class="sxs-lookup"><span data-stu-id="895a8-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="895a8-225">Рассмотрим сценарий, в котором эта вкладка должна быть видна, и только при условии, что в случае активации диаграммы (на листе книги Excel по умолчанию).</span><span class="sxs-lookup"><span data-stu-id="895a8-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="895a8-226">Начните с назначения обработчиков.</span><span class="sxs-lookup"><span data-stu-id="895a8-226">Begin by assigning handlers.</span></span> <span data-ttu-id="895a8-227">Это обычно делается в `Office.onReady` методе, как в следующем примере, в котором обработчики (созданные на последующих шагах) назначаются `onActivated` `onDeactivated` событиям и событиям всех диаграмм на листе.</span><span class="sxs-lookup"><span data-stu-id="895a8-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

<span data-ttu-id="895a8-228">Затем определите обработчики.</span><span class="sxs-lookup"><span data-stu-id="895a8-228">Next, define the handlers.</span></span> <span data-ttu-id="895a8-229">Ниже приведен простой пример `showDataTab` , но в этой статье описывается [обработка ошибки хострестартнидед](#handling-the-hostrestartneeded-error) далее в этой статье, чтобы получить более надежную версию функции.</span><span class="sxs-lookup"><span data-stu-id="895a8-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handling-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="895a8-230">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="895a8-230">About this code, note:</span></span>

- <span data-ttu-id="895a8-231">Office определяет время обновления состояния ленты.</span><span class="sxs-lookup"><span data-stu-id="895a8-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="895a8-232">Метод  [Office. Ribbon. рекуеступдате](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) ставит в очередь запрос на обновление.</span><span class="sxs-lookup"><span data-stu-id="895a8-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="895a8-233">Метод будет разрешать `Promise` объект, как только он помещается в очередь запроса, а не при фактическом обновлении ленты.</span><span class="sxs-lookup"><span data-stu-id="895a8-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="895a8-234">Параметр для `requestUpdate` метода — это объект [риббонупдатердата](/javascript/api/office/office.ribbonupdaterdata) , который (1) указывает на вкладку по ее идентификатору точно так же, *как указано в JSON* , а (2) указывает на видимость вкладки.</span><span class="sxs-lookup"><span data-stu-id="895a8-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="895a8-235">При наличии нескольких настраиваемых контекстных вкладок, которые должны быть видны в одном контексте, просто добавьте в массив дополнительные объекты Tab `tabs` .</span><span class="sxs-lookup"><span data-stu-id="895a8-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="895a8-236">Обработчик для скрытия вкладки почти идентичен, за исключением того, что свойство возвращает значение `visible` `false` .</span><span class="sxs-lookup"><span data-stu-id="895a8-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="895a8-237">Библиотека JavaScript для Office также предоставляет несколько интерфейсов (типов) для упрощения создания `RibbonUpdateData` объекта.</span><span class="sxs-lookup"><span data-stu-id="895a8-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="895a8-238">Ниже приведена `showDataTab` функция TypeScript, которая использует эти типы.</span><span class="sxs-lookup"><span data-stu-id="895a8-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="895a8-239">Переключать видимость вкладок и включенное состояние кнопки одновременно</span><span class="sxs-lookup"><span data-stu-id="895a8-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="895a8-240">Этот `requestUpdate` метод также используется для переключения состояния включенного или отключенного настраиваемой кнопки на настраиваемой контекстной вкладке или на пользовательской вкладке "основной". Подробнее об этом можно узнать в статье [Включение и отключение команд надстроек](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="895a8-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="895a8-241">Возможны сценарии, в которых необходимо одновременно изменить видимость вкладки и включенное состояние кнопки.</span><span class="sxs-lookup"><span data-stu-id="895a8-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="895a8-242">Это можно сделать с помощью одного вызова `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="895a8-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="895a8-243">Ниже приведен пример, в котором одновременно становится доступна кнопка на основной вкладке, в которой отображается контекстная вкладка.</span><span class="sxs-lookup"><span data-stu-id="895a8-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

<span data-ttu-id="895a8-244">В следующем примере кнопка включена на той же контекстной вкладке, которая становится видимой.</span><span class="sxs-lookup"><span data-stu-id="895a8-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="localizing-the-json-blob"></a><span data-ttu-id="895a8-245">Локализация большого двоичного объекта JSON</span><span class="sxs-lookup"><span data-stu-id="895a8-245">Localizing the JSON blob</span></span>

<span data-ttu-id="895a8-246">Передаваемый большой двоичный объект JSON `requestCreateControls` не локализуется точно так же, как разметка манифеста для настраиваемых основных вкладок (описывается в разделе [Локализация из манифеста](../develop/localization.md#control-localization-from-the-manifest)).</span><span class="sxs-lookup"><span data-stu-id="895a8-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="895a8-247">Вместо этого локализация должна выполняться во время выполнения с использованием различных больших двоичных объектов JSON для каждого языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="895a8-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="895a8-248">Мы рекомендуем использовать `switch` оператор, который проверяет свойство [Office. Context. displayLanguage](/javascript/api/office/office.context#displayLanguage) .</span><span class="sxs-lookup"><span data-stu-id="895a8-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="895a8-249">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="895a8-249">The following is an example:</span></span>

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
                          "label": "Data",
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
                          "label": "Données",
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

<span data-ttu-id="895a8-250">Затем код вызывает функцию для получения локализованного объекта BLOB, который передается `requestCreateControls` , как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="895a8-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a><span data-ttu-id="895a8-251">Обработка ошибки Хострестартнидед</span><span class="sxs-lookup"><span data-stu-id="895a8-251">Handling the HostRestartNeeded error</span></span>

<span data-ttu-id="895a8-252">В некоторых случаях Office не может обновить ленту и возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="895a8-252">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="895a8-253">Например, если после обновления у надстройки другой набор настраиваемых команд, приложение Office необходимо закрыть и снова открыть.</span><span class="sxs-lookup"><span data-stu-id="895a8-253">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="895a8-254">Пока это действие не будет выполнено, метод `requestUpdate` будет возвращать ошибку `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="895a8-254">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="895a8-255">Ниже приведен пример обработки этой ошибки.</span><span class="sxs-lookup"><span data-stu-id="895a8-255">The following is an example of how to handle this error.</span></span> <span data-ttu-id="895a8-256">В этом случае метод `reportError` выводит сообщение об ошибке для пользователя.</span><span class="sxs-lookup"><span data-stu-id="895a8-256">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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
