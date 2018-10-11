# <a name="scopes-element"></a><span data-ttu-id="414af-101">Элемент Scopes</span><span class="sxs-lookup"><span data-stu-id="414af-101">Scopes element</span></span>

<span data-ttu-id="414af-102">Содержит разрешения, необходимые надстройке для работы с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="414af-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="414af-103">Магазин Office использует элемент Scopes для создания диалогового окна подтверждения.</span><span class="sxs-lookup"><span data-stu-id="414af-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="414af-104">Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="414af-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="414af-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="414af-105">Child elements</span></span>

|  <span data-ttu-id="414af-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="414af-106">Element</span></span> |  <span data-ttu-id="414af-107">Тип</span><span class="sxs-lookup"><span data-stu-id="414af-107">Type</span></span>  |  <span data-ttu-id="414af-108">Описание</span><span class="sxs-lookup"><span data-stu-id="414af-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="414af-109">**Scope**</span><span class="sxs-lookup"><span data-stu-id="414af-109">**Scope**</span></span>                |  <span data-ttu-id="414af-110">string</span><span class="sxs-lookup"><span data-stu-id="414af-110">string</span></span>     |   <span data-ttu-id="414af-111">Имя разрешения на доступ к Microsoft Graph (например, Files.Read.All).</span><span class="sxs-lookup"><span data-stu-id="414af-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="414af-112">Пример</span><span class="sxs-lookup"><span data-stu-id="414af-112">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
