---
title: Как определить правильный порядок элементов манифеста
description: Узнайте, как определить правильный порядок расположения дочерних элементов в родительском элементе.
ms.date: 08/15/2019
localization_priority: Normal
ms.openlocfilehash: 68eaa21af9a86c910a9e8701ef2ea6cf95a64000
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477763"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>Как определить правильный порядок элементов манифеста

XML-элементы в манифесте надстройки Office должны располагаться под правильным родительском элементом *и* в определенном порядке относительно друг друга под родительским элементом.

Нужный порядок указывается в XSD-файлах в папке [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). XSD-файлы упорядочены в подпапках для области задач, контента и почтовых надстроек.

Например, в элементе `<OfficeApp>` элементы `<Id>`, `<Version>`, `<ProviderName>` должны располагаться в указанном порядке. Если добавляется элемент `<AlternateId>`, он должен размещаться между элементами `<Id>` и `<Version>`. Ваш манифест будет недопустимым и надстройка не загрузится, если любой из элементов находится в неправильном порядке.

> [!NOTE]
> [Средство проверки в Office-ADDIN-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) использует то же самое сообщение об ошибке, если элемент находится в неупорядоченном виде, если элемент находится в неправильном родительском элементе. В сообщении об ошибке указывается, что для родительского элемента этот дочерний элемент не является допустимым. Если появляется такая ошибка, но при этом в справочной документации указано, что дочерний элемент *является* допустимым для родительского, значит проблема вероятно связана с тем, что дочерний элемент помещен в неправильном порядке.

В следующих разделах показаны элементы манифеста в том порядке, в котором они должны отображаться. Существуют небольшие отличия в зависимости от того, `type` имеет ли атрибут `<OfficeApp>` элемента значение `TaskPaneApp`, `ContentApp`или `MailApp`. Чтобы эти разделы не стали слишком громоздкими, строго сложный `<VersionOverrides>` элемент разбивается на отдельные разделы.

> [!Note]
> Не все отображаемые элементы являются обязательными. Если `minOccurs` значение элемента равно **0** в [схеме](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), элемент является необязательным.

## <a name="basic-task-pane-add-in-element-ordering"></a>Упорядочение элементов базовой области задач

```
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
```

\*Рассмотрите [сортировку элементов надстройки области задач в VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) для упорядочивания дочерних элементов VersionOverrides.

## <a name="basic-mail-add-in-element-ordering"></a>Упорядочение элементов базовой почтовой надстройки

```
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

\*В статье упорядочение [элементов почтовых ящиков в VersionOverrides ver. 1,0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) и почтовых почтовых [элементов надстройки в VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для упорядочивания дочерних элементов VersionOverrides.

## <a name="basic-content-add-in-element-ordering"></a>Упорядочение элементов базовой надстройки контента

```
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
    <VersionOverrides>
```

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>Упорядочение элементов надстройки области задач в VersionOverrides

```
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
      <Hosts>
        <Host>
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>Упорядочение элементов почтовой надстройки в VersionOverrides ver. 1.0

```
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

\*Объект VersionOverrides со `type` значением `VersionOverridesV1_1`, а не `VersionOverridesV1_0`, может быть вложен в конце внешнего VersionOverrides. Сведения о порядке элементов почтовых ящиков в [VersionOverrides ver. 1,1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) для упорядочивания элементов `VersionOverridesV1_1`в.

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>Упорядочение элементов почтовой надстройки в VersionOverrides ver. 1.1

```
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

## <a name="see-also"></a>См. также

- [Справочник по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md)
