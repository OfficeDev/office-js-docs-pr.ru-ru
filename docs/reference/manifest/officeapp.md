---
title: Элемент OfficeApp в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433168"
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

|**Element**|**Контентная надстройка**|**Почтовая надстройка**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[Description](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Permissions](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>Может содержать:

|**Element**|**Контентная надстройка**|**Почтовая надстройка**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requirements](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Permissions](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Dictionary](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)||x||

## <a name="attributes"></a>Атрибуты

|||
|:-----|:-----|
|xmlns|Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.|
|xmlns:xsi|Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.|
|xsi:type|Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`|
