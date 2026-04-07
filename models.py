from pydantic import BaseModel, Field


class DetailItem(BaseModel):
    model_config = {'extra': 'forbid'}

    part_number: str = Field(alias='КаталожныйНомер', min_length=1)
    quantity: float = Field(alias='Количество', ge=0)
    upd: str = Field(alias='УПД', min_length=1)

