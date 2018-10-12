# <a name="requirements-element"></a>Элемент Requirements

Указывает минимальный набор требований API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.

**Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Элемент, в котором содержится

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|**Элемент**|**Контентные**|**Почтовые**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Наборы](sets.md)|x|x|x|
|[Методы](methods.md)|x||x|

## <a name="remarks"></a>Замечания

Дополнительные сведения о наборах требований см. в статье [версии и наборы требований Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

