---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 770c764db6d8d7d1d2e870e48437de7c8f887101
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641461"
---
# <a name="officeapp-element"></a>Элемент OfficeApp

Корневой элемент в манифесте надстройки Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>Содержится в

 _none_

## <a name="must-contain"></a>Должен содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Id](id.md)|x|x|x|
|[Версия](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[Описание](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Разрешения](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requirements](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Разрешения](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Dictionary](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|x|x|x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|xmlns|Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.|
|xmlns:xsi|Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.|
|xsi:type|Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`|
