using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using ERP_NEW.BLL.DTO.SelectedDTO;
using ERP_NEW.BLL.DTO.ModelsDTO;
using ERP_NEW.BLL.Interfaces;
using ERP_NEW.DAL.Entities.Models;
using ERP_NEW.DAL.Entities.QueryModels;
using ERP_NEW.DAL.Interfaces;
using FirebirdSql.Data.FirebirdClient;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace ERP_NEW.BLL.Services
{
    public class EmployeesService : IEmployeesService
    {
        private IUnitOfWork Database { get; set; }
        private IRepository<EmployeesInfo> employeesInfo;
        private IRepository<Employees> employees;
        private IRepository<EmployeesDetails> employeesDetails;
        private IRepository<EmployeePhoto> employeePhoto;

        private IMapper mapper;
        


        public EmployeesService(IUnitOfWork uow)
        {
            Database = uow;
            employeesInfo = Database.GetRepository<EmployeesInfo>();
            employees = Database.GetRepository<Employees>();
            employeesDetails = Database.GetRepository<EmployeesDetails>();
            employeePhoto = Database.GetRepository<EmployeePhoto>();

            var config = new MapperConfiguration(cfg =>
             {
                 cfg.CreateMap<EmployeesInfo, EmployeesInfoDTO>();
               
                 cfg.CreateMap<EmployeePhoto, EmployeePhotoDTO>();
              

             });

            mapper = config.CreateMapper();
        }

        public IEnumerable<EmployeesInfoDTO> GetEmployeeHistory(decimal employeeNumber)
        {
            FbParameter[] Parameters =
            {
                new FbParameter("Number", employeeNumber),
            };
            
            string procName = @"select * from ""GetEmployeeHistory""(@Number)";

            return mapper.Map<IEnumerable<EmployeesInfo>, List<EmployeesInfoDTO>>(employeesInfo.SQLExecuteProc(procName, Parameters));
        }

        public IEnumerable<EmployeesInfoDTO> GetEmployeesWorking()
        {
            string procName = @"select * from ""GetEmployeesWorking""";

            return mapper.Map<IEnumerable<EmployeesInfo>, List<EmployeesInfoDTO>>(employeesInfo.SQLExecuteProc(procName));
        }

        public IEnumerable<EmployeesInfoDTO> GetEmployeesNotWorking()
        {
            string procName = @"select * from ""GetEmployeesNotWorking""";

            return mapper.Map<IEnumerable<EmployeesInfo>, List<EmployeesInfoDTO>>(employeesInfo.SQLExecuteProc(procName));
        }

        public IEnumerable<EmployeesInfoDTO> GetEmployeesWorkingByDeparmentId(int departmentId)
        {
            string procName = @"select * from ""GetEmployeesWorking""";

            return mapper.Map<IEnumerable<EmployeesInfo>, List<EmployeesInfoDTO>>(employeesInfo.SQLExecuteProc(procName).Where(bdsm => bdsm.DepartmentID == departmentId));
        }
     
        public EmployeePhotoDTO GetPhotoById(int EmployeesId)
        {
            FbParameter[] Parameters =
                {
                    new FbParameter("EmployeeId", EmployeesId)
                };

            string procName = @"select * from ""GetEmployeePhotoById""(@EmployeeId)";

            return (mapper.Map<IEnumerable<EmployeePhoto>, List<EmployeePhotoDTO>>(employeePhoto.SQLExecuteProc(procName, Parameters))).SingleOrDefault();
        }        
        public void Dispose()
        {
            Database.Dispose();
        }
    }
}
