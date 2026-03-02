--Thang lương
select *from PA_SALARY_SCALE
-- Ngạch lương
select *from PA_SALARY_GRADE where SALARY_SCALE_ID in(select ID from PA_SALARY_SCALE)
-- Bậc lương
select *from PA_SALARY_LEVEL where SALARY_GRADE_ID in(select ID from PA_SALARY_GRADE where SALARY_SCALE_ID in(select ID from PA_SALARY_SCALE))


select *from HU_WORKING


