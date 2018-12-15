# <a name="appdomain-element"></a>Элемент AppDomain

Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Примечания

Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).
