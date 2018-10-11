# <a name="appdomains-element"></a>Элемент AppDomains

Определяет все домены, кроме указанного в элементе SourceLocation, которые надстройка Office будет использовать для загрузки страниц. Для каждого дополнительного домена укажите элемент AppDomain.

 **Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>Элемент, в котором содержится

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

[AppDomain](appdomain.md)

## <a name="remarks"></a>Замечания

По умолчанию надстройка может загружать страницы из домена, указанного в элементе **SourceLocation**. Для загрузки страниц из других доменов, укажите домены в элементах **AppDomains** и **AppDomain**. Этот элемент не может быть пустым. 
