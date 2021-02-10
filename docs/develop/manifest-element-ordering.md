---
title: Как определить правильный порядок элементов манифеста
description: Узнайте, как определить правильный порядок расположения дочерних элементов в родительском элементе.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ee80167a76861209e814dc6c272720feb3a9cf1
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173915"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="8d9d2-103">Как определить правильный порядок элементов манифеста</span><span class="sxs-lookup"><span data-stu-id="8d9d2-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="8d9d2-104">XML-элементы в манифесте надстройки Office должны располагаться под правильным родительском элементом *и* в определенном порядке относительно друг друга под родительским элементом.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="8d9d2-105">Нужный порядок указывается в XSD-файлах в папке [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="8d9d2-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="8d9d2-106">XSD-файлы упорядочены в подпапках для области задач, контента и почтовых надстроек.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="8d9d2-107">Например, в элементе `<OfficeApp>` элементы `<Id>`, `<Version>`, `<ProviderName>` должны располагаться в указанном порядке.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="8d9d2-108">Если добавляется элемент `<AlternateId>`, он должен размещаться между элементами `<Id>` и `<Version>`.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="8d9d2-109">Ваш манифест будет недопустимым и надстройка не загрузится, если любой из элементов находится в неправильном порядке.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9d2-110">В [манифесте office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) используется то же сообщение об ошибке, что и при неправильном родительском элементе.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="8d9d2-111">В сообщении об ошибке указывается, что для родительского элемента этот дочерний элемент не является допустимым.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="8d9d2-112">Если появляется такая ошибка, но при этом в справочной документации указано, что дочерний элемент *является* допустимым для родительского, значит проблема вероятно связана с тем, что дочерний элемент помещен в неправильном порядке.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="8d9d2-113">В следующих разделах появятся элементы манифеста в том порядке, в котором они должны отображаться.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="8d9d2-114">Существуют различия в зависимости от того, является ли атрибут `type` `<OfficeApp>` элемента , или `TaskPaneApp` `ContentApp` `MailApp` .</span><span class="sxs-lookup"><span data-stu-id="8d9d2-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="8d9d2-115">Чтобы эти разделы не стали слишком громоздкими, сложный элемент разбивается `<VersionOverrides>` на отдельные разделы.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="8d9d2-116">Не все показанные элементы являются обязательными.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="8d9d2-117">Если значение `minOccurs` элемента в схеме [](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3) **— 0,** элемент является необязательным.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="8d9d2-118">Упорядочение элементов надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="8d9d2-118">Basic task pane add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="TaskPaneApp">
    <Id>
    <AlternateID>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
        <Sets>
            <Set>
        <Methods>
            <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <Permissions>
    <Dictionary>
        <TargetDialects>
        <QueryUri>
        <CitationText>
        <DictionaryName>
        <DictionaryHomePage>
    <VersionOverrides>*
    <ExtendedOverrides>
```

<span data-ttu-id="8d9d2-119">\*См. порядок элементов надстройки области задач в [VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) для упорядочения элементов VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="8d9d2-120">Базовый порядок элементов почтовой надстройки</span><span class="sxs-lookup"><span data-stu-id="8d9d2-120">Basic mail add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="MailApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <FormSettings>
        <Form>
        <DesktopSettings>
            <SourceLocation>
            <RequestedHeight>
        <TabletSettings>
            <SourceLocation>
            <RequestedHeight>
        <PhoneSettings>
            <SourceLocation>
    <Permissions>
    <Rule>
    <DisableEntityHighlighting>
    <VersionOverrides>*
```

<span data-ttu-id="8d9d2-121">\*См. порядок элементов почтовой надстройки в [versionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) и Mail в [versionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для упорядочения элементов VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="8d9d2-122">Упорядочение элементов базовой контентной надстройки</span><span class="sxs-lookup"><span data-stu-id="8d9d2-122">Basic content add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="ContentApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl >
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <Methods>
        <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <RequestedWidth>
    <RequestedHeight>
    <Permissions>
    <AllowSnapshot>
    <VersionOverrides>*
```

<span data-ttu-id="8d9d2-123">\*См. [порядок элементов контентной надстройки в VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) для упорядочения элементов VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="8d9d2-124">Порядок элементов надстройки области задач в VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8d9d2-124">Task pane add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <Runtimes>
                <Runtime>
            <AllFormFactors>
                <ExtensionPoint>
                    <Script>
                        <SourceLocation>
                    <Page>
                        <SourceLocation>
                    <Metadata>
                        <SourceLocation>
                    <Namespace>
            <DesktopFormFactor>
                <GetStarted>
                    <Title>
                    <Description>
                    <LearnMoreUrl>
                <FunctionFile>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                    <CustomTab>
                        <OverriddenByRibbonApi>
                        <Group> (can be below <ControlGroup>)
                            <OverriddenByRibbonApi>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <OverriddenByRibbonApi>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Enabled>
                                <Items>
                                    <Item>
                                        <OverriddenByRibbonApi>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <ControlGroup> (can be above <Group>)
                        <Label>
                        <InsertAfter> (or <InsertBefore>)
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
        <Resources>
            <Images>
                <Image>
                    <Override>
            <Urls>
                <Url>
                    <Override>
            <ShortStrings>
                <String>
                    <Override>
            <LongStrings>
                <String>
                    <Override>
        <WebApplicationInfo>
            <Id>
            <MsaId>
            <Resource>
            <Scopes>
                <Scope>
            <Authorizations>
                <Authorization>
                    <Resource>
                    <Scopes>
                        <Scope>
        <EquivalentAddins>
            <EquivalentAddin>
                <ProgId>
                <DisplayName>
                <FileName>
                <Type>
```

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="8d9d2-125">Упорядочение элементов почтовой надстройки в VersionOverrides Ver.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="8d9d2-126">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9d2-126">1.0</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <DesktopFormFactor>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>
                                <Action>
                                    <SourceLocation>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Items>
                                    <Item>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <VersionOverrides>*
```

<span data-ttu-id="8d9d2-127">\* VersionOverrides со значением, а не , может быть вложен в `type` `VersionOverridesV1_1` конце `VersionOverridesV1_0` внешнего VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="8d9d2-128">См. порядок элементов почтовой надстройки в [versionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для упорядочения элементов в `VersionOverridesV1_1` .</span><span class="sxs-lookup"><span data-stu-id="8d9d2-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="8d9d2-129">Упорядочение элементов почтовой надстройки в VersionOverrides Ver.</span><span class="sxs-lookup"><span data-stu-id="8d9d2-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="8d9d2-130">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9d2-130">1.1</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
    <Sets>
        <Set>
    <Hosts>
    <Host>
        <DesktopFormFactor>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <SourceLocation>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                    <SourceLocation>
                <SourceLocation>
                <Label>
                <CommandSurface>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="8d9d2-131">Порядок элементов контентной надстройки в VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8d9d2-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="8d9d2-132">См. также</span><span class="sxs-lookup"><span data-stu-id="8d9d2-132">See also</span></span>

- [<span data-ttu-id="8d9d2-133">Справочник по манифестам надстройки Office (1.1)</span><span class="sxs-lookup"><span data-stu-id="8d9d2-133">Reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="8d9d2-134">Официальные определения схемы</span><span class="sxs-lookup"><span data-stu-id="8d9d2-134">Official schema definitions</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
