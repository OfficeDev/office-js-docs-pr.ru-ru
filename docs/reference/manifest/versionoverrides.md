---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация элемента VersionOverrides для Office файлов манифеста надстройок (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 657bdebbc88993badd9d0e60946239edd55d5533
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042149"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте. Его детская разметка может переопределять часть разметки в базовом манифесте (или в родительской **ВерсииOverrides).** **VersionOverrides** — это детский элемент корневого элемента [OfficeApp](officeapp.md) в манифесте или родительского **элемента VersionOverrides.** Этот элемент поддерживается в схеме манифеста v1.1 и более поздней версии, но определяется в отдельных схемах VersionOverrides.

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен VersionOverrides схемы. Допустимые значения различаются в зависимости от `<VersionOverrides>` **значения xsi:type** этого элемента и **значения xsi:type** родительского `<OfficeApp>` элемента. Ниже [см. значения Пространства](#namespace-values) имен.|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространства имен

Ниже перечислены необходимое значение атрибута **xmlns** в зависимости от **значения xsi:type** корневого `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны быть `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны быть `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому значение **xmlns** зависит от значения `<VersionOverrides>` **xsi:type** этого элемента:
  - Когда **xsi:type** `VersionOverridesV1_0` , **xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides` .
  - Когда **xsi:type** `VersionOverridesV1_1` , **xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> В настоящее время Outlook 2016 поддерживает схему VersionOverrides v1.1 и `VersionOverridesV1_1` тип.

## <a name="variant-schemas"></a>Схемы вариантов

Для каждого из возможных значений **xmlns** существует отдельная схема, поэтому каждая из них имеет отдельную справочную страницу.

- [VersionOverrides 1.0 TaskPane](versionoverrides-1-0-taskpane.md)
- [VersionOverrides 1.0 Content](versionoverrides-1-0-content.md)
- [VersionOverrides 1.0 Mail](versionoverrides-1-0-mail.md)
- [VersionOverrides 1.1 Mail](versionoverrides-1-1-mail.md)
