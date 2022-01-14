---
title: VersionOverrides элемент 1.0 в файле манифеста для надстройки контента
description: Справочная документация элемента VersionOverrides (контент) для Office файлов манифеста надстройок (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a9cd431f0e8fb4a7abe49103522e04900d9bcfd
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042190"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>VersionOverrides элемент 1.0 в файле манифеста для надстройки контента

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с обзором элемента [VersionOverrides,](versionoverrides.md)который содержит важную информацию о атрибутах и вариантах элемента.

## <a name="child-elements"></a>Дочерние элементы

В следующей таблице используется только версия 1.0 элементов **VersionOverrides** и только надстройки контента.

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Нет  | В настоящее время не может быть в VersionOverrides 1.0 для надстройок контента. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |

## <a name="example"></a>Пример

Ниже приведен простой пример. Более полные примеры см. в манифестах для примеров надстройок в Office примерах [кода надстройки.](https://github.com/OfficeDev/PnP-OfficeAddins)

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
