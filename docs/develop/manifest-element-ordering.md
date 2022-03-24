---
title: Как определить правильный порядок элементов манифеста
description: Узнайте, как определить правильный порядок расположения дочерних элементов в родительском элементе.
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2bd29b4df28ae3d4831da625580097f1970a532d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743400"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Как определить правильный порядок элементов манифеста

XML-элементы в манифесте надстройки Office должны располагаться под правильным родительском элементом *и* в определенном порядке относительно друг друга под родительским элементом.

Нужный порядок указывается в XSD-файлах в папке [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). XSD-файлы упорядочены в подпапках для области задач, контента и почтовых надстроек.

Например, в элементе `<OfficeApp>` элементы `<Id>`, `<Version>`, `<ProviderName>` должны располагаться в указанном порядке. Если добавляется элемент `<AlternateId>`, он должен размещаться между элементами `<Id>` и `<Version>`. Ваш манифест будет недопустимым и надстройка не загрузится, если любой из элементов находится в неправильном порядке.

> [!NOTE]
> [Валидатор в office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) использует то же сообщение об ошибке, когда элемент не в порядке, как и при неправильном родительском элементе. В сообщении об ошибке указывается, что для родительского элемента этот дочерний элемент не является допустимым. Если появляется такая ошибка, но при этом в справочной документации указано, что дочерний элемент *является* допустимым для родительского, значит проблема вероятно связана с тем, что дочерний элемент помещен в неправильном порядке.

В следующих разделах отображаются элементы манифеста в порядке, в котором они должны отображаться. Существуют различия в зависимости от `type` `<OfficeApp>` `TaskPaneApp`того, является ли атрибут элемента , или . `ContentApp``MailApp` Чтобы эти разделы не стали слишком громоздкими, `<VersionOverrides>` очень сложный элемент разбивается на отдельные разделы.

> [!Note]
> Не все из показанных элементов являются обязательными. Если значение `minOccurs` элемента **0** в [схеме](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), элемент необязателен.

## <a name="basic-task-pane-add-in-element-ordering"></a>Базовое упорядочение элементов надстройки области задач

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

\*См [. элемент надстройки](#task-pane-add-in-element-ordering-within-versionoverrides) области задач, упорядочение в VersionOverrides для заказа детских элементов VersionOverrides.

## <a name="basic-mail-add-in-element-ordering"></a>Базовый заказ элементов надстройки почты

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

\*См. элемент Mail [add-in ordering in VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) и [Mail add-in ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для заказа детских элементов VersionOverrides.

## <a name="basic-content-add-in-element-ordering"></a>Упорядочение элементов базовой надстройки контента

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

\*См [. элемент надстройки контента, упорядочение в VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) для заказа детских элементов VersionOverrides.

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>Заказ элементов надстройки области задач в VersionOverrides

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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>Заказ элемента надстройки почты в ВерсииOverrides Ver. 1.0

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

\* VersionOverrides со значением `type` `VersionOverridesV1_1`, а `VersionOverridesV1_0`не , может быть вложен в конце внешних VersionOverrides. См. в рублях элемент mail [add-in ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для заказа элементов в `VersionOverridesV1_1`.

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>Заказ элемента надстройки почты в ВерсииOverrides Ver. 1.1

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

## <a name="content-add-in-element-ordering-within-versionoverrides"></a>Упорядочение элемента надстройки контента в VersionOverrides

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a>См. также

- [Ссылка на манифесты Office надстройки (v1.1)](../develop/add-in-manifests.md)
- [Определения официальной схемы](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
