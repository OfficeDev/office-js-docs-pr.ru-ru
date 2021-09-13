---
title: Элемент AppDomain в файле манифеста
description: Указывает дополнительные домены, которые используются вашей надстройки и должны доверяться Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: c17195e6d9d3f4f22465c8aa1fc626afd3eb06c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153729"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Указывает дополнительный домен, Office должен доверять, в дополнение к домену, указанному в [элементе SourceLocation.](sourcelocation.md) Указание домена имеет такие эффекты:

- Он позволяет открывать страницы, маршруты или другие ресурсы в домене непосредственно в корневой области задач надстройки на настольных Office платформах. (Указание домена в **AppDomain** не требуется для Office в Интернете или для открытия ресурса в IFrame, равно как и для открытия ресурса в диалоговом диалоговом окрашиваемом [API](../../develop/dialog-api-in-office-add-ins.md)диалогов .)
- Он позволяет страницам в домене Office.js API-вызовы из IFrames в надстройки.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Если для домена имеется явный порт, включай его (например, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Если нужно доверять поддомену, включай его (например, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). Subdomain `mysubdomain.mydomain.com` и `mydomain.com` являются различными доменами. Если обоим нужно доверять, то оба должны быть в отдельных **элементах AppDomain.**
> 4. Перечисление того же домена, что и домен, указанный в [элементе SourceLocation,](sourcelocation.md) не имеет эффекта и может вводить в заблуждение. В частности, при разработке не требуется создавать элемент `localhost` **AppDomain** для `localhost` .
> 5. Не включайте в домен какие-либо сегменты URL-адреса. Например, не включайте полный URL-адрес страницы.
> 6. Не *помещай* закрываю черту "/", на значение.

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).
