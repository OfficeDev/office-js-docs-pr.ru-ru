---
title: Элемент authorization в файле манифеста
description: Указывает внешний ресурс, на который веб-приложению надстройки требуется авторизация и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294264"
---
# <a name="authorization-element"></a><span data-ttu-id="655c8-103">Элемент authorization</span><span class="sxs-lookup"><span data-stu-id="655c8-103">Authorization element</span></span>

<span data-ttu-id="655c8-104">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="655c8-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="655c8-105">**Авторизация** является дочерним элементом элемента [authorizations](authorizations.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="655c8-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="655c8-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="655c8-106">Child elements</span></span>

|  <span data-ttu-id="655c8-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="655c8-107">Element</span></span> |  <span data-ttu-id="655c8-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="655c8-108">Required</span></span>  |  <span data-ttu-id="655c8-109">Описание</span><span class="sxs-lookup"><span data-stu-id="655c8-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="655c8-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="655c8-110">**Resource**</span></span>  |  <span data-ttu-id="655c8-111">Да</span><span class="sxs-lookup"><span data-stu-id="655c8-111">Yes</span></span>   |  <span data-ttu-id="655c8-112">Задает URL-адрес внешнего ресурса.</span><span class="sxs-lookup"><span data-stu-id="655c8-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="655c8-113">Scopes</span><span class="sxs-lookup"><span data-stu-id="655c8-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="655c8-114">Да</span><span class="sxs-lookup"><span data-stu-id="655c8-114">Yes</span></span>  |  <span data-ttu-id="655c8-115">Задает разрешения, необходимые надстройке для ресурса.</span><span class="sxs-lookup"><span data-stu-id="655c8-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="655c8-116">Пример</span><span class="sxs-lookup"><span data-stu-id="655c8-116">Example</span></span>

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
