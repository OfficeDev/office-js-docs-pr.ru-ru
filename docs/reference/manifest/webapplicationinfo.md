# <a name="webapplicationinfo-element"></a>Элемент WebApplicationInfo

Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:

- *Ресурс* OAuth 2.0, для которого могут потребоваться разрешения ведущему приложению Office.
- *Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.

**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.  

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Id**    |  Да   |  **Идентификатор приложения** связанный с надстройкой службы, зарегистрированной в конечной точке Azure Active Directory 2.0.|
|  **Ресурс**  |  Да   |  Указывает **URI идентификатора приложения** надстройки, зарегистрированной в конечной точке Azure Active Directory 2.0.|
|  [Области применения](scopes.md)                |  Нет  |  Указывает разрешения, необходимые надстройке для работы с Microsoft Graph.  |

> [!NOTE] 
> В настоящее время необходимо, чтобы ресурс надстройки соответствовал ее узлу. Office запрашивает маркер для надстройки, только если может подтвердить право собственности. В настоящее время для этого необходимо, чтобы надстройка размещалась под полным доменным именем ресурса.

## <a name="webapplicationinfo-example"></a>Пример WebApplicationInfo

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
