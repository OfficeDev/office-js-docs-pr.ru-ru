---
title: Элемент authorization в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477945"
---
# <a name="authorization-element"></a><span data-ttu-id="d4540-102">Элемент authorization</span><span class="sxs-lookup"><span data-stu-id="d4540-102">Authorization element</span></span>

<span data-ttu-id="d4540-103">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="d4540-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="d4540-104">**Авторизация** является дочерним элементом [](authorizations.md) элемента authorizations в манифесте.</span><span class="sxs-lookup"><span data-stu-id="d4540-104">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d4540-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d4540-105">Child elements</span></span>

|  <span data-ttu-id="d4540-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="d4540-106">Element</span></span> |  <span data-ttu-id="d4540-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d4540-107">Required</span></span>  |  <span data-ttu-id="d4540-108">Описание</span><span class="sxs-lookup"><span data-stu-id="d4540-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d4540-109">**Resource**</span><span class="sxs-lookup"><span data-stu-id="d4540-109">**Resource**</span></span>  |  <span data-ttu-id="d4540-110">Да</span><span class="sxs-lookup"><span data-stu-id="d4540-110">Yes</span></span>   |  <span data-ttu-id="d4540-111">Задает URL-адрес внешнего ресурса.</span><span class="sxs-lookup"><span data-stu-id="d4540-111">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="d4540-112">Scopes</span><span class="sxs-lookup"><span data-stu-id="d4540-112">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="d4540-113">Да</span><span class="sxs-lookup"><span data-stu-id="d4540-113">Yes</span></span>  |  <span data-ttu-id="d4540-114">Задает разрешения, необходимые надстройке для ресурса.</span><span class="sxs-lookup"><span data-stu-id="d4540-114">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="d4540-115">Пример</span><span class="sxs-lookup"><span data-stu-id="d4540-115">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
