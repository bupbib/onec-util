from pydantic import BaseModel, Field


class DetailItem(BaseModel):
    model_config = {'extra': 'forbid'}
    
    part_number: str = Field(alias='КаталожныйНомер')
    quantity: float = Field(alias='Количество')
    upd: str = Field(alias='УПД')

