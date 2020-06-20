---
title: Элемент AppDomain в файле манифеста
description: Указывает дополнительные домены, используемые надстройкой, и которые должны быть доверенными для Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778650"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Задает дополнительный домен, который должен быть доверенным для Office, в дополнение к тому, что указано в [элементе SourceLocation](sourcelocation.md). Указание домена включает в себя следующие эффекты:

- Он позволяет открывать страницы, маршруты и другие ресурсы в домене непосредственно в корневой области задач на настольных платформах Office. (Указание домена в **домене AppDomain** не требуется для Office в Интернете или открытие ресурса в iframe, а также не требуется для открытия ресурса в диалоговом окне, открываемом с помощью [API диалогового окна](../../develop/dialog-api-in-office-add-ins.md).)
- Он позволяет страницам в домене совершать Office.js вызовы API из IFrames в надстройке.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Если для домена существует явный порт, включите его (например, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).
> 3. Если дочерний домен должен быть доверенным, включите его (например, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ). Дочерний домен `mysubdomain.mydomain.com` и `mydomain.com` разные домены. Если необходимо, чтобы оба были доверенными, они должны находиться в отдельных элементах **AppDomain** .
> 4. Перечисление того же домена, что и в [элементе SourceLocation](sourcelocation.md) , не оказывает никакого действия и может привести к некоторому определению. В частности, когда вы разрабатываете `localhost` , вам не нужно создавать элемент **AppDomain** для `localhost` .
> 5. Не включайте ни один из сегментов URL-адреса за пределами домена. Например, не включайте полный URL-адрес страницы.
> 6. *Не* ставьте закрывающую косую черту ("/") для значения.

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).
