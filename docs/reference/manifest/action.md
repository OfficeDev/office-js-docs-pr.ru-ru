---
title: Элемент Action в файле манифеста
description: Этот элемент указывает действие, выполняемое при выборе пользователем элемента управления "Кнопка" или "меню".
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: c542cec38b400100014c51c978c8fcd71a546f2a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608805"
---
# <a name="action-element"></a><span data-ttu-id="9a200-103">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="9a200-103">Action element</span></span>

<span data-ttu-id="9a200-104">Задает действие, выполняемое при выборе пользователем элемента управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) ".</span><span class="sxs-lookup"><span data-stu-id="9a200-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="9a200-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9a200-105">Attributes</span></span>

|  <span data-ttu-id="9a200-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="9a200-106">Attribute</span></span>  |  <span data-ttu-id="9a200-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9a200-107">Required</span></span>  |  <span data-ttu-id="9a200-108">Описание</span><span class="sxs-lookup"><span data-stu-id="9a200-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9a200-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9a200-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9a200-110">Да</span><span class="sxs-lookup"><span data-stu-id="9a200-110">Yes</span></span>  | <span data-ttu-id="9a200-111">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="9a200-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="9a200-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9a200-112">Child elements</span></span>

|  <span data-ttu-id="9a200-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="9a200-113">Element</span></span> |  <span data-ttu-id="9a200-114">Описание</span><span class="sxs-lookup"><span data-stu-id="9a200-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9a200-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9a200-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="9a200-116">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="9a200-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="9a200-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9a200-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="9a200-118">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9a200-118">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="9a200-119"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="9a200-119"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="9a200-120">Определяет идентификатор для контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="9a200-120">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="9a200-121"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="9a200-121"> [Title](#title)</span></span> | <span data-ttu-id="9a200-122">Определяет заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="9a200-122">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="9a200-123"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="9a200-123"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="9a200-124">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="9a200-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="9a200-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9a200-125">xsi:type</span></span>

<span data-ttu-id="9a200-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="9a200-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="9a200-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="9a200-128">FunctionName</span></span>

<span data-ttu-id="9a200-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="9a200-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="9a200-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9a200-132">SourceLocation</span></span>

<span data-ttu-id="9a200-133">Обязательный элемент, если **xsi: Type** — "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="9a200-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9a200-134">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9a200-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="9a200-135">Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="9a200-135">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="9a200-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="9a200-136">TaskpaneId</span></span>

<span data-ttu-id="9a200-137">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="9a200-137">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9a200-138">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="9a200-138">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="9a200-139">Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="9a200-139">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="9a200-140">Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область.</span><span class="sxs-lookup"><span data-stu-id="9a200-140">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="9a200-141">Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но содержимое области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.</span><span class="sxs-lookup"><span data-stu-id="9a200-141">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="9a200-142">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a200-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="9a200-143">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="9a200-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="9a200-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="9a200-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="9a200-146">Должность</span><span class="sxs-lookup"><span data-stu-id="9a200-146">Title</span></span>

<span data-ttu-id="9a200-147">Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="9a200-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9a200-148">Определяет заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="9a200-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="9a200-149">В приведенном ниже примере показано действие, в котором используется элемент **Title** .</span><span class="sxs-lookup"><span data-stu-id="9a200-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="9a200-150">Обратите внимание, что **заголовок** не назначается строке напрямую.</span><span class="sxs-lookup"><span data-stu-id="9a200-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="9a200-151">Вместо этого ему назначается идентификатор ресурса (Resid), который определяется в разделе **Resources (ресурсы** ) манифеста.</span><span class="sxs-lookup"><span data-stu-id="9a200-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="9a200-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="9a200-152">SupportsPinning</span></span>

<span data-ttu-id="9a200-153">Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane.</span><span class="sxs-lookup"><span data-stu-id="9a200-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="9a200-154">Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="9a200-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="9a200-155">Включите этот элемент со значением `true` для поддержки закрепления области задач.</span><span class="sxs-lookup"><span data-stu-id="9a200-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="9a200-156">Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента.</span><span class="sxs-lookup"><span data-stu-id="9a200-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="9a200-157">Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="9a200-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9a200-158">Хотя `SupportsPinning` элемент был введен в [наборе требований 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), в настоящее время он поддерживается только для подписчиков Office 365 с помощью следующих компонентов.</span><span class="sxs-lookup"><span data-stu-id="9a200-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Office 365 subscribers using the following.</span></span>
> - <span data-ttu-id="9a200-159">Outlook 2016 или более поздняя версия в Windows (сборка 7628,1000 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="9a200-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="9a200-160">Outlook 2016 или более поздней версии в Mac (сборка 16.13.503 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="9a200-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="9a200-161">Современная версия Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="9a200-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
