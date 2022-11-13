using Core7ExcelPdfExample.Dtos.Categories;

namespace Core7ExcelPdfExample.Business;

public interface ICategoryService
{
    List<CategoryListDto> CategoryList();
}
