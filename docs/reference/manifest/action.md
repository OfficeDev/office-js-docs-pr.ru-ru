# <a name="action-element"></a><span data-ttu-id="f92d9-101">Элемент Action</span><span class="sxs-lookup"><span data-stu-id="f92d9-101">Action element</span></span>

<span data-ttu-id="f92d9-102">Указывает действие, которое необходимо выполнить, когда пользователь выбирает элемент управления [Кнопка](control.md#button-control) или [Меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="f92d9-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>
 
## <a name="attributes"></a><span data-ttu-id="f92d9-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f92d9-103">Attributes</span></span>

|  <span data-ttu-id="f92d9-104">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f92d9-104">Attribute</span></span>  |  <span data-ttu-id="f92d9-105">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f92d9-105">Required</span></span>  |  <span data-ttu-id="f92d9-106">Описание</span><span class="sxs-lookup"><span data-stu-id="f92d9-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f92d9-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f92d9-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="f92d9-108">Да</span><span class="sxs-lookup"><span data-stu-id="f92d9-108">Yes</span></span>  | <span data-ttu-id="f92d9-109">Тип выполняемого действия</span><span class="sxs-lookup"><span data-stu-id="f92d9-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="f92d9-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f92d9-110">Child elements</span></span>

|  <span data-ttu-id="f92d9-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="f92d9-111">Element</span></span> |  <span data-ttu-id="f92d9-112">Описание</span><span class="sxs-lookup"><span data-stu-id="f92d9-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="f92d9-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="f92d9-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="f92d9-114">Указывает имя выполняемой функции.</span><span class="sxs-lookup"><span data-stu-id="f92d9-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="f92d9-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f92d9-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="f92d9-116">Указывает расположение исходного файла для этого действия.</span><span class="sxs-lookup"><span data-stu-id="f92d9-116">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="f92d9-117">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="f92d9-117">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="f92d9-118">Определяет идентификатор контейнера области задач.</span><span class="sxs-lookup"><span data-stu-id="f92d9-118">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="f92d9-119">Title</span><span class="sxs-lookup"><span data-stu-id="f92d9-119">Title</span></span>](#title) | <span data-ttu-id="f92d9-120">Определяет настраиваемый заголовок области задач.</span><span class="sxs-lookup"><span data-stu-id="f92d9-120">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="f92d9-121">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="f92d9-121">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="f92d9-122">Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).</span><span class="sxs-lookup"><span data-stu-id="f92d9-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="f92d9-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f92d9-123">xsi:type</span></span>

<span data-ttu-id="f92d9-p101">Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="f92d9-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="f92d9-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="f92d9-126">FunctionName</span></span>

<span data-ttu-id="f92d9-p102">Обязательный элемент, если атрибуту **xsi:type** присвоено значение "ExecuteFunction". Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="f92d9-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="f92d9-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f92d9-130">SourceLocation</span></span>

<span data-ttu-id="f92d9-p103">Обязательный элемент, если атрибуту **xsi:type** присвоено значение "ShowTaskpane". Указывает расположение исходного файла для этого действия. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="f92d9-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="f92d9-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="f92d9-134">TaskpaneId</span></span>

<span data-ttu-id="f92d9-p104">Необязательный элемент, когда для атрибута **xsi:type** задано значение "ShowTaskpane". Определяет идентификатор для контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но оглавление области будет заменено соответствующим дочерним элементом "SourceLocation" элемента Action.</span><span class="sxs-lookup"><span data-stu-id="f92d9-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span> 

> [!NOTE]
> <span data-ttu-id="f92d9-140">Этот элемент не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="f92d9-140">Note: This element is not supported in Outlook.</span></span>

<span data-ttu-id="f92d9-141">В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="f92d9-141">The following example shows two actions that share the same **TaskpaneId**.</span></span> 

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

<span data-ttu-id="f92d9-p105">В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="f92d9-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="f92d9-144">Title</span><span class="sxs-lookup"><span data-stu-id="f92d9-144">Title</span></span>
<span data-ttu-id="f92d9-p106">Необязательный элемент, когда для атрибута **xsi:type** задано значение "ShowTaskpane". Определяет настраиваемый заголовок области задач для этого действия.</span><span class="sxs-lookup"><span data-stu-id="f92d9-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action.</span></span> 

<span data-ttu-id="f92d9-147">В приведенных ниже примерах показаны два действия, для которых используется элемент **Title**.</span><span class="sxs-lookup"><span data-stu-id="f92d9-147">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
``` 

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
``` 

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
``` 

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
``` 

## <a name="supportspinning"></a><span data-ttu-id="f92d9-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="f92d9-148">SupportsPinning</span></span>

<span data-ttu-id="f92d9-p107">Необязательный элемент, когда для атрибута **xsi:type** задано значение "ShowTaskpane". Содержащие элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`. Включите этот элемент со значением `true` для поддержки закрепления области задач. Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента. Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="f92d9-p107">Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="f92d9-154">Примечание. В настоящее время элемент SupportsPinning поддерживается только в Outlook 2016 для Windows (сборка 7628.1000 или более поздней версии).</span><span class="sxs-lookup"><span data-stu-id="f92d9-154">Note: SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


