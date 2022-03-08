---
title: VersionOverrides элемент 1.0 в файле манифеста для надстройки контента
description: Справочная документация элемента VersionOverrides (контент) для Office файлов манифеста надстройок (XML).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ef083ef5df322c230292625576e36db8923d00c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341053"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>VersionOverrides элемент 1.0 в файле манифеста для надстройки контента

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с обзором элемента [VersionOverrides](versionoverrides.md), который содержит важную информацию о атрибутах и вариантах элемента.

## <a name="child-elements"></a>Дочерние элементы

В следующей таблице используется только версия 1.0 элементов **VersionOverrides** и только надстройки контента.

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Нет  | В настоящее время не может быть в VersionOverrides 1.0 для надстройок контента. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |

## <a name="example"></a>Пример

Ниже приведен простой пример. Более сложные примеры см. в манифестах для примерных надстройок в Office [примерах кода надстройки](https://github.com/OfficeDev/PnP-OfficeAddins).

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
