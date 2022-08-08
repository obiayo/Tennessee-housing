select * from dbo.Sheet1$

select convert (date,saledate)
from dbo.Sheet1$

select a.ParcelID, a.PropertyAddress, b.ParcelID, b.PropertyAddress, isnull (a.PropertyAddress, b.PropertyAddress)
from dbo.Sheet1$ a
join dbo.Sheet1$ b
	on a.ParcelID = b.ParcelID
	and a.[UniqueID ] <> b.[UniqueID ]
where a.PropertyAddress is null

update a
set PropertyAddress = isnull (a.PropertyAddress, b.PropertyAddress)
from dbo.Sheet1$ a
join dbo.Sheet1$ b
	on a.ParcelID = b.ParcelID
	and a.[UniqueID ] <> b.[UniqueID ]
where a.PropertyAddress is null

Select 
Substring(PropertyAddress,1,Charindex(',', PropertyAddress) -1) as Address, 
Substring(PropertyAddress, Charindex(',', PropertyAddress) + 1, len(PropertyAddress)) as Address 
from dbo.Sheet1$

Alter table Sheet1$
Add PropertySplitAddress Nvarchar(255);

Update Sheet1$
Set PropertySplitAddress= Substring(PropertyAddress, 1, Charindex(',', PropertyAddress) -1)

Alter table Sheet1$
add PropertySplitCity Nvarchar(255);

Update Sheet1$
Set PropertySplitCity = Substring(PropertyAddress, Charindex(',', PropertyAddress) + 1, len(PropertyAddress))

select * from dbo.Sheet1$

select
PARSENAME (replace(owneraddress, ',', '.') ,3)
,PARSENAME (replace(owneraddress, ',', '.') ,2)
,PARSENAME (replace(owneraddress, ',', '.') ,1)
from dbo.Sheet1$




Alter table Sheet1$
add OwnerSplitAddress Nvarchar(255);

Update Sheet1$
Set OwnerSplitAddress = PARSENAME (replace(owneraddress, ',', '.') ,3)

Alter table Sheet1$
add OwnerSplitCity Nvarchar(255);

Update Sheet1$
Set OwnerSplitCity = PARSENAME (replace(owneraddress, ',', '.') ,2)

Alter table Sheet1$
add OwnerSplitState Nvarchar(255);

Update Sheet1$
Set OwnerSplitState = PARSENAME (replace(owneraddress, ',', '.') ,1)

select * from dbo.Sheet1$


select distinct(SoldAsVacant), Count(SoldAsVacant)
from dbo.Sheet1$
group by SoldAsVacant
order by 2

Select SoldAsVacant
 CASE When SoldAsVacant='Y' THEN 'Yes'
	  When SoldAsVacant = 'N' THEN 'No'
	  ELSE SoldAsVacant
	  END
From dbo.Sheet1$

Update Sheet1$
SET SoldAsVacant = CASE When SoldAsVacant = 'Y' THEN 'Yes'
	When SoldAsVacant = 'N' THEN 'NO'
	ELSE SoldAsVacant
	END

WITH RowNumCTE as(
Select *,
	Row_number() over (
	Partition by ParcelID,
				 PropertyAddress,
				 SalePrice,
				 SaleDate,
				 LegalReference
				 Order by
					UniqueID
					) row_num
from sheet1$
)
Delete
From RowNumCTE 
where row_num >1
--order by PropertyAddress

select *
from Sheet1$

alter table sheet1$
drop column SaleDate, OwnerAddress, TaxDistrict, PropertyAddress

