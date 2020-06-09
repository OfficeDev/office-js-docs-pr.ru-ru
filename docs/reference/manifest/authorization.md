---
title: Элемент authorization в файле манифеста
description: Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8c6249706b8eef11f579378fe5c9dc83016d17c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608763"
---
# <a name="authorization-element"></a><span data-ttu-id="07a7e-103">Элемент authorization</span><span class="sxs-lookup"><span data-stu-id="07a7e-103">Authorization element</span></span>

<span data-ttu-id="07a7e-104">Указывает внешние ресурсы, к которым веб-приложению надстройки требуется авторизация, и необходимые разрешения.</span><span class="sxs-lookup"><span data-stu-id="07a7e-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="07a7e-105">**Авторизация** является дочерним элементом элемента [authorizations](authorizations.md) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="07a7e-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="07a7e-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="07a7e-106">Child elements</span></span>

|  <span data-ttu-id="07a7e-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="07a7e-107">Element</span></span> |  <span data-ttu-id="07a7e-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="07a7e-108">Required</span></span>  |  <span data-ttu-id="07a7e-109">Описание</span><span class="sxs-lookup"><span data-stu-id="07a7e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="07a7e-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="07a7e-110">**Resource**</span></span>  |  <span data-ttu-id="07a7e-111">Да</span><span class="sxs-lookup"><span data-stu-id="07a7e-111">Yes</span></span>   |  <span data-ttu-id="07a7e-112">Задает URL-адрес внешнего ресурса.</span><span class="sxs-lookup"><span data-stu-id="07a7e-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="07a7e-113">Scopes</span><span class="sxs-lookup"><span data-stu-id="07a7e-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="07a7e-114">Да</span><span class="sxs-lookup"><span data-stu-id="07a7e-114">Yes</span></span>  |  <span data-ttu-id="07a7e-115">Задает разрешения, необходимые надстройке для ресурса.</span><span class="sxs-lookup"><span data-stu-id="07a7e-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="07a7e-116">Пример</span><span class="sxs-lookup"><span data-stu-id="07a7e-116">Example</span></span>

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
