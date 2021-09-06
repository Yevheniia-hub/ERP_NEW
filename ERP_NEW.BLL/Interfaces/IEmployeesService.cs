using System.Collections.Generic;
using ERP_NEW.BLL.DTO.ModelsDTO;
using ERP_NEW.BLL.DTO.SelectedDTO;
using System;


namespace ERP_NEW.BLL.Interfaces
{
    public interface IEmployeesService
    {
        IEnumerable<EmployeesInfoDTO> GetEmployeeHistory(decimal employeeNumber);
        IEnumerable<EmployeesInfoDTO> GetEmployeesWorking();
        IEnumerable<EmployeesInfoDTO> GetEmployeesNotWorking();
        IEnumerable<EmployeesInfoOnlyWithWeldStampDTO> GetEmployeesWorkingWithWeldStamp();
        IEnumerable<EmployeesInfoNonPhotoDTO> GetEmployeesWorkingNonPhoto();
        EmployeePhotoDTO GetPhotoById(int EmployeesId);
        IEnumerable<DepartmentsDTO> GetDepartments();
        IEnumerable<EmployeesInfoDTO> GetEmployeesWorkingByDeparmentId(int departmentId);
        IEnumerable<EmployeeVisitScheduleDTO> GetEmployeeVisitScheduleProc(int employeeId, DateTime startDate, DateTime endDate);
        //IEnumerable<EmployeesDetailsDTO> GetEmployesDetals();


        void Dispose();
    }
}
