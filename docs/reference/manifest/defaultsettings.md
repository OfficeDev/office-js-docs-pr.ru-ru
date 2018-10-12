# <a name="defaultsettings-element"></a>Элемент DefaultSettings

Указывает исходное расположение по умолчанию и другие параметры по умолчанию для контентной надстройки или надстройки области задач.

**Тип надстройки:** контентные надстройки и надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Элемент, в котором содержится

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|**Элемент**|**Контентные**|**Почтовые**|**Области задач**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Замечания

Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие параметры по умолчанию с помощью элемента [FormSettings](formsettings.md).

