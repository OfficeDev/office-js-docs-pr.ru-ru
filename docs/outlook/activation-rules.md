---
title: Правила активации для надстроек Outlook
description: Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: a5fc107c27feb5b0535727a42b4d56d21f7dcbc4
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076814"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="4ba6b-103">Правила активации контекстных надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="4ba6b-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="4ba6b-p101">Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки. Это верно для всех надстроек, для которых используется схема манифеста 1.1. Затем пользователь может выбрать надстройку из пользовательского интерфейса Outlook, чтобы запустить ее для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="4ba6b-107">На следующем изображении показаны надстройки Outlook, активируемые в области надстройки для сообщения в области чтения.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![Панели приложений с активированными приложениями для чтения почты.](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="4ba6b-109">Указание правил активации в манифесте</span><span class="sxs-lookup"><span data-stu-id="4ba6b-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="4ba6b-110">Чтобы Outlook активировать надстройки для определенных условий, укажите правила активации в манифесте надстройки с помощью одного из `Rule` следующих элементов:</span><span class="sxs-lookup"><span data-stu-id="4ba6b-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements:</span></span>

- <span data-ttu-id="4ba6b-111">[элемент Rule (MailApp complexType)](../reference/manifest/rule.md), задающий отдельное правило;</span><span class="sxs-lookup"><span data-stu-id="4ba6b-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="4ba6b-112">[Элемент Rule (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection), совмещающий несколько правил с помощью логических операторов.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>
    

 > [!NOTE]
 > <span data-ttu-id="4ba6b-113">Элемент, `Rule` который используется для указания отдельного [](../reference/manifest/rule.md) правила, имеет сложный абстрактный тип Правила.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="4ba6b-114">Каждый из следующих типов правил расширяет этот абстрактный `Rule` сложный тип.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="4ba6b-115">Следовательно, указывая отдельное правило в манифесте, необходимо использовать атрибут [xsi:type](https://www.w3.org/TR/xmlschema-1/), чтобы определить один из перечисленных ниже типов правил.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="4ba6b-116">Например, следующее правило определяет правило[ItemIs](../reference/manifest/rule.md#itemis-rule): `<Rule xsi:type="ItemIs" ItemType="Message" />`</span><span class="sxs-lookup"><span data-stu-id="4ba6b-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule: `<Rule xsi:type="ItemIs" ItemType="Message" />`</span></span>
 > 
 > <span data-ttu-id="4ba6b-117">Атрибут применяется к правилам активации `FormType` в манифесте v1.1, но не определен `VersionOverrides` в v1.0.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="4ba6b-118">Поэтому его нельзя использовать, когда [ItemIs](../reference/manifest/rule.md#itemis-rule) используется в `VersionOverrides` узле.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="4ba6b-p104">В таблице ниже перечислены доступные типы элементов. Дополнительные сведения см. под таблицей и в статьях, перечисленных в статье [Создание надстроек Outlook для форм чтения](read-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p104">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="4ba6b-121">**Имя правила**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-121">**Rule name**</span></span>|<span data-ttu-id="4ba6b-122">**Применимые формы**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-122">**Applicable forms**</span></span>|<span data-ttu-id="4ba6b-123">**Описание**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="4ba6b-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="4ba6b-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="4ba6b-125">Чтение, создание</span><span class="sxs-lookup"><span data-stu-id="4ba6b-125">Read, Compose</span></span>|<span data-ttu-id="4ba6b-p105">Проверяет, относится ли текущий элемент к определенному типу (сообщение или встреча). Кроме того, оно может проверять класс элемента, тип формы и, при необходимости, класс сообщения элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p105">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="4ba6b-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="4ba6b-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="4ba6b-129">Чтение</span><span class="sxs-lookup"><span data-stu-id="4ba6b-129">Read</span></span>|<span data-ttu-id="4ba6b-130">Проверяет, содержит ли выделенный элемент вложение.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="4ba6b-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="4ba6b-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="4ba6b-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="4ba6b-132">Read</span></span>|<span data-ttu-id="4ba6b-p106">Проверяет, содержит ли выделенный элемент одну или несколько известных сущностей. Дополнительные сведения см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p106">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="4ba6b-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="4ba6b-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="4ba6b-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="4ba6b-136">Read</span></span>|<span data-ttu-id="4ba6b-137">Проверяет, содержит ли адрес электронной почты отправителя, тема и/или тело выбранного элемента совпадение с регулярным выражением. Подробнее: [Использование регулярных правил активации выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="4ba6b-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="4ba6b-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="4ba6b-139">Чтение, создание</span><span class="sxs-lookup"><span data-stu-id="4ba6b-139">Read, Compose</span></span>|<span data-ttu-id="4ba6b-140">Объединяет набор правил, чтобы можно было создавать более сложные правила.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="4ba6b-141">Правило ItemIs</span><span class="sxs-lookup"><span data-stu-id="4ba6b-141">ItemIs rule</span></span>

<span data-ttu-id="4ba6b-142">Сложный тип **ItemIs** определяет правило, которое имеет значение **true**, если текущий элемент совпадает с типом элемента и (необязательно) с классом сообщения элемента (если он указан в правиле).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="4ba6b-143">Укажите один из следующих типов элементов в `ItemType` атрибуте правила **ItemIs.**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="4ba6b-144">В манифесте можно указать несколько правил **ItemIs**.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="4ba6b-145">Значение simpleType атрибута ItemType определяет типы элементов Outlook, поддерживающих надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="4ba6b-146">**Value**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-146">**Value**</span></span>|<span data-ttu-id="4ba6b-147">**Описание**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="4ba6b-148">**Встреча**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-148">**Appointment**</span></span>|<span data-ttu-id="4ba6b-p108">Указывает элемент в календаре Outlook. Это может быть элемент собрания, для которого был отправлен ответ и у которого есть организатор и участники, или встреча без организатора или участника, которая просто представляет собой элемент календаря. Соответствует классу сообщений IPM.Appointment в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p108">Specifies an item in an Outlook calendar. This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="4ba6b-151">**Сообщение**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-151">**Message**</span></span>|<span data-ttu-id="4ba6b-152">Указывает один из приведенных ниже элементов, обычно поступающих в папку "Входящие".</span><span class="sxs-lookup"><span data-stu-id="4ba6b-152">Specifies one of the following items received in typically the Inbox:</span></span> <ul><li><p><span data-ttu-id="4ba6b-p109">Сообщение электронной почты. Соответствует классу сообщений IPM.Note в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p109">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="4ba6b-p110">Запрос на собрание, ответ или отклонение. Соответствует следующим классам сообщений в Outlook:</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p110">A meeting request, response, or cancellation. This corresponds to the following  message classes in Outlook:</span></span></p><p><span data-ttu-id="4ba6b-157">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="4ba6b-157">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="4ba6b-158">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="4ba6b-158">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="4ba6b-159">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="4ba6b-159">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="4ba6b-160">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="4ba6b-160">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="4ba6b-161">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="4ba6b-161">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="4ba6b-162">Атрибут используется для указания режима (чтения или составить), в котором должна активироваться `FormType` надстройка.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-162">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="4ba6b-163">Атрибут ItemIs определяется в схеме v1.1 и более поздней, но `FormType` не `VersionOverrides` в v1.0.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-163">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="4ba6b-164">Не включайте атрибут при определении команд `FormType` надстройки.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-164">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="4ba6b-165">После активации надстройки можно использовать свойство [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) для получения элемента, выбранного в текущий момент в Outlook, и свойство [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для получения типа текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-165">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="4ba6b-166">Атрибут можно дополнительно использовать для указания класса сообщений элемента, а атрибут указывает, должно ли правило быть верным, если элемент является подклассом `ItemClass` `IncludeSubClasses` указанного класса. </span><span class="sxs-lookup"><span data-stu-id="4ba6b-166">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="4ba6b-167">Дополнительные сведения о классах сообщений см. в статье [Типы элементов и классы сообщений](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-167">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="4ba6b-168">В приведенном ниже примере показано правило **ItemIs**, которое отображает надстройку на панели надстройки Outlook, когда пользователь просматривает сообщение.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-168">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message:</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="4ba6b-169">В приведенном ниже примере показано правило **ItemIs**, которое отображает надстройку на панели надстройки Outlook, когда пользователь просматривает сообщение или встречу.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="4ba6b-170">Правило ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="4ba6b-170">ItemHasAttachment rule</span></span>


<span data-ttu-id="4ba6b-171">Сложный тип определяет правило, которое проверяет, содержит ли выбранный элемент `ItemHasAttachment` вложение.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-171">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="4ba6b-172">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="4ba6b-172">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="4ba6b-173">Перед тем как элемент становится доступным надстройке, сервер проверяет, содержат ли тема и основной текст строку, которая с высокой вероятностью может быть одной из известных сущностей.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-173">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities.</span></span> <span data-ttu-id="4ba6b-174">Если какие-либо из этих сущностями найдены, он помещается в коллекцию известных сущностями, к которые можно получить доступ с помощью или метода `getEntities` `getEntitiesByType` этого элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-174">If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="4ba6b-175">Вы можете указать правило, используя которое отображает вашу надстройка, когда в элементе присутствует объект `ItemHasKnownEntity` указанного типа.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-175">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item.</span></span> <span data-ttu-id="4ba6b-176">В атрибуте правила можно указать следующие известные `EntityType` `ItemHasKnownEntity` сущности:</span><span class="sxs-lookup"><span data-stu-id="4ba6b-176">You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule:</span></span>

- <span data-ttu-id="4ba6b-177">Address</span><span class="sxs-lookup"><span data-stu-id="4ba6b-177">Address</span></span>
- <span data-ttu-id="4ba6b-178">Contact</span><span class="sxs-lookup"><span data-stu-id="4ba6b-178">Contact</span></span>
- <span data-ttu-id="4ba6b-179">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="4ba6b-179">EmailAddress</span></span>
- <span data-ttu-id="4ba6b-180">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4ba6b-180">MeetingSuggestion</span></span>
- <span data-ttu-id="4ba6b-181">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4ba6b-181">PhoneNumber</span></span>
- <span data-ttu-id="4ba6b-182">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4ba6b-182">TaskSuggestion</span></span>
- <span data-ttu-id="4ba6b-183">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="4ba6b-183">URL</span></span>
    
<span data-ttu-id="4ba6b-184">Можно дополнительно включить регулярное выражение в атрибут, чтобы надстройка была показана только в том случае, если объект соответствует обычному выражению `RegularExpression` в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-184">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present.</span></span> <span data-ttu-id="4ba6b-185">Чтобы получить совпадения с регулярными выражениями, указанными в правилах, можно использовать метод или метод для выбранного Outlook `ItemHasKnownEntity` `getRegExMatches` `getFilteredEntitiesByName` элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-185">To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="4ba6b-186">В следующем примере показана коллекция элементов, отображающих надстройки, когда в сообщении присутствует одно из указанных известных `Rule` сущностями.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-186">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="4ba6b-187">В следующем примере показано правило с атрибутом, который активирует надстройку, если в сообщении присутствует URL-адрес, содержащий слово `ItemHasKnownEntity` `RegularExpression` "contoso".</span><span class="sxs-lookup"><span data-stu-id="4ba6b-187">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="4ba6b-188">Дополнительные сведения о сущностях в правилах активации см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="4ba6b-188">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="4ba6b-189">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="4ba6b-189">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="4ba6b-190">Сложный тип определяет правило, которое использует регулярное выражение, чтобы соответствовать содержимому указанного `ItemHasRegularExpressionMatch` свойства элемента.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-190">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item.</span></span> <span data-ttu-id="4ba6b-191">Если текст, соответствующий регулярному выражению, обнаруживается в заданном свойстве элемента, Outlook активирует панель надстроек и отображает надстройку.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-191">If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in.</span></span> <span data-ttu-id="4ba6b-192">Для получения совпадений для указанного регулярного выражения можно использовать объект или метод объекта, который представляет выбранный в настоящее время `getRegExMatches` `getRegExMatchesByName` элемент.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-192">You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="4ba6b-193">В следующем примере показана надстройка, активируемая, когда тело выбранного элемента содержит `ItemHasRegularExpressionMatch` "яблоко", "банан" или "кокос", игнорируя случай.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-193">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="4ba6b-194">Дополнительные сведения об использовании правила см. в статью Использование правил активации регулярных выражений, чтобы показать Outlook `ItemHasRegularExpressionMatch` [надстройку.](use-regular-expressions-to-show-an-outlook-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="4ba6b-194">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="4ba6b-195">Правило RuleCollection</span><span class="sxs-lookup"><span data-stu-id="4ba6b-195">RuleCollection rule</span></span>


<span data-ttu-id="4ba6b-196">Сложный `RuleCollection` тип объединяет несколько правил в одно правило.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-196">The `RuleCollection` complex type combines multiple rules into a single rule.</span></span> <span data-ttu-id="4ba6b-197">Можно указать, следует ли сочетать правила в коллекции с логическим или логическим и с помощью `Mode` атрибута.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-197">You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="4ba6b-p117">Если указан логический оператор "И", то для отображения надстройки элемент должен соответствовать всем заданным правилам в коллекции. Если выбран логический оператор "ИЛИ", надстройка будет отображаться при наличии элемента, соответствующего любому из заданных правил в коллекции.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p117">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="4ba6b-200">Вы можете объединить `RuleCollection` правила для формирования сложных правил.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-200">You can combine `RuleCollection` rules to form complex rules.</span></span> <span data-ttu-id="4ba6b-201">Следующий пример активирует надстройку, если пользователь просматривает встречу или сообщение, а тема или основной текст сообщения или встречи содержит адрес.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-201">The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="4ba6b-202">Следующий пример активирует надстройку, если пользователь создает сообщение или просматривает встречу, а тема или основной текст встречи содержит адрес.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-202">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="4ba6b-203">Ограничения для правил и регулярных выражений</span><span class="sxs-lookup"><span data-stu-id="4ba6b-203">Limits for rules and regular expressions</span></span>


<span data-ttu-id="4ba6b-204">Чтобы обеспечить удобство использования надстроек Outlook, следует соблюдать рекомендации по активации и использованию API-интерфейсов.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-204">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines.</span></span> <span data-ttu-id="4ba6b-205">В следующей таблице показаны общие ограничения для регулярных выражений и правил, но существуют определенные правила для различных приложений.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-205">The following table shows general limits for regular expressions and rules but there are specific rules for different applications.</span></span> <span data-ttu-id="4ba6b-206">Дополнительные сведения см. в дополнительных сведениях о ограничениях для активации и [API JavaScript](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) для Outlook надстройки и устранения неполадок Outlook [надстройки.](troubleshoot-outlook-add-in-activation.md)</span><span class="sxs-lookup"><span data-stu-id="4ba6b-206">For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="4ba6b-207">**Элемент надстройки**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-207">**Add-in element**</span></span>|<span data-ttu-id="4ba6b-208">**Рекомендации**</span><span class="sxs-lookup"><span data-stu-id="4ba6b-208">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="4ba6b-209">Размер манифеста</span><span class="sxs-lookup"><span data-stu-id="4ba6b-209">Manifest Size</span></span>|<span data-ttu-id="4ba6b-210">Не более 256 КБ.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-210">No larger than 256 KB.</span></span>|
|<span data-ttu-id="4ba6b-211">Правила</span><span class="sxs-lookup"><span data-stu-id="4ba6b-211">Rules</span></span>|<span data-ttu-id="4ba6b-212">Не более 15 правил.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-212">No more than 15 rules.</span></span>|
|<span data-ttu-id="4ba6b-213">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="4ba6b-213">ItemHasKnownEntity</span></span>|<span data-ttu-id="4ba6b-214">Полнофункциональный клиент Outlook применит правило к первому мегабайту основного текста, но не будет применять его к остальному тексту.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-214">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="4ba6b-215">Регулярные выражения</span><span class="sxs-lookup"><span data-stu-id="4ba6b-215">Regular Expressions</span></span>|<span data-ttu-id="4ba6b-216">Правила ItemHasKnownEntity или ItemHasRegularExpressionMatch для всех Outlook приложений:</span><span class="sxs-lookup"><span data-stu-id="4ba6b-216">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="4ba6b-p120">Задавайте не более 5 регулярных выражений в правилах активации для надстройки Outlook. Если этот предел будет превышен, установить надстройку будет невозможно.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-p120">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="4ba6b-219">Задавайте регулярные выражения, ожидаемые результаты которых возвращаются в первых 50 совпадениях с помощью метода <b>getRegExMatches</b>.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-219">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="4ba6b-220">Указывайте в регулярных выражениях утверждения с просмотром вперед, а не утверждения с просмотром назад `(?<=text)` или отрицательные утверждения с просмотром назад `(?<!text)`.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-220">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="4ba6b-221">Задавайте регулярные выражения, соответствия которым не превышают ограничений, указанных в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="4ba6b-221">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="4ba6b-222">Ограничение длины для результата, соответствующего регулярному выражению</span><span class="sxs-lookup"><span data-stu-id="4ba6b-222">Limit on length of a regex match</span></span></th><th><span data-ttu-id="4ba6b-223">Полнофункциональные клиенты Outlook</span><span class="sxs-lookup"><span data-stu-id="4ba6b-223">Outlook rich clients</span></span></th><th><span data-ttu-id="4ba6b-224">Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="4ba6b-224">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="4ba6b-225">Основной текст элемента в виде простого текста</span><span class="sxs-lookup"><span data-stu-id="4ba6b-225">Item body is plain text</span></span></td><td><span data-ttu-id="4ba6b-226">1,5 КБ</span><span class="sxs-lookup"><span data-stu-id="4ba6b-226">1.5 KB</span></span></td><td><span data-ttu-id="4ba6b-227">3 КБ</span><span class="sxs-lookup"><span data-stu-id="4ba6b-227">3 KB</span></span></td></tr><tr><td><span data-ttu-id="4ba6b-228">Основной текст элемента в виде HTML-кода</span><span class="sxs-lookup"><span data-stu-id="4ba6b-228">Item body it HTML</span></span></td><td><span data-ttu-id="4ba6b-229">3 КБ</span><span class="sxs-lookup"><span data-stu-id="4ba6b-229">3 KB</span></span></td><td><span data-ttu-id="4ba6b-230">3 КБ</span><span class="sxs-lookup"><span data-stu-id="4ba6b-230">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="4ba6b-231">См. также</span><span class="sxs-lookup"><span data-stu-id="4ba6b-231">See also</span></span>

- [<span data-ttu-id="4ba6b-232">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="4ba6b-232">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="4ba6b-233">Ограничения активации и API JavaScript для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="4ba6b-233">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="4ba6b-234">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4ba6b-234">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="4ba6b-235">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="4ba6b-235">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
