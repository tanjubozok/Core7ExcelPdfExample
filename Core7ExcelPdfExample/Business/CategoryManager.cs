using Core7ExcelPdfExample.Data;
using Core7ExcelPdfExample.Dtos.Categories;

namespace Core7ExcelPdfExample.Business;

public class CategoryManager : ICategoryService
{
    private readonly DatabaseContext _databaseContext;

    public CategoryManager(DatabaseContext databaseContext)
    {
        _databaseContext = databaseContext;
    }

    public List<CategoryListDto> CategoryList()
    {
        var categoryList = _databaseContext.Categories.ToList();

        List<CategoryListDto> categoryListDto = new();
        foreach (var item in categoryList)
        {
            CategoryListDto categoryDto = new()
            {
                CategoryID = item.CategoryID,
                CategoryName = item.CategoryName,
                Description = item.Description
            };
            categoryListDto.Add(categoryDto);
        }
        return categoryListDto;
    }
}
