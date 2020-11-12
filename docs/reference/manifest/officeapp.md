---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996321"
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
|[екстендедоверридес](extendedoverrides.md)|||x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|xmlns|Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.|
|xmlns:xsi|Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.|
|xsi:type|Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`|
