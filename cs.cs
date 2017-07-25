using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using EaseSource.AnDa.SMT.Web.DbContext;
using EaseSource.AnDa.SMT.Web.MvcExtensions;
using EaseSource.AnDa.SMT.Entity;
using Microsoft.AspNetCore.Authorization;
using EaseSource.AnDa.SMT.Web.Utility;
using Microsoft.AspNetCore.Identity;
using EaseSource.AnDa.SMT.Web.ViewModels;
using AutoMapper;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using System.Security.Claims;
using EaseSource.Dingtalk.Entity;
using EaseSource.Dingtalk.Interfaces;
using EaseSource.AnDa.SMT.Web.ViewModels.Response;
using System.IO;
using EaseSource.AnDa.SMT.Web.Services;
using Microsoft.Extensions.Logging;

namespace EaseSource.AnDa.SMT.Web.Controllers
{
    /// <summary>
    /// ��������API
    /// </summary>
    [Authorize]
    [Route("api/[controller]")]
    public class WorkOrderController : SMTControllerBase
    {
        private readonly IMapper mapper;
        private readonly IPolishedBundleSpecSheetService bundleSpecSheetSVC;
        private readonly ILogger<WorkOrderController> logger;
        private readonly IBundleInStockSpecSheetService bundleInStockSpecSheetSVC;

        public WorkOrderController(AnDaDbContext context, UserManager<AnDaUser> userManager, IMapper mapper, IDingtalkMessenger dtMessenger, IPolishedBundleSpecSheetService bundleSpecSheetSVC, ILogger<WorkOrderController> logger, IBundleInStockSpecSheetService bundleInStockSpecSheetSVC)
        : base(context, userManager, dtMessenger)
        {
            this.mapper = mapper;
            this.bundleSpecSheetSVC = bundleSpecSheetSVC;
            this.logger = logger;
            this.bundleInStockSpecSheetSVC = bundleInStockSpecSheetSVC;
        }

        // GET: api/WorkOrder/GetAll?pageSize=10&pageNo=1
        /// <summary>
        /// ��ȡ���������������ṩ��ҳ����
        /// </summary>
        /// <param name="pageSize">ÿҳ���������ݣ��粻�ṩ��Ĭ��ÿҳ20��</param>
        /// <param name="pageNo">ҳ�ţ��粻�ṩ��Ĭ�ϵ�1ҳ</param>
        /// <param name="blockNumber">��ѡ�����ϱ�ţ�������ṩ�����ʾ���л���</param>
        /// <param name="orderNumber">��ѡ��������ţ�������ṩ�����ʾ���й���</param>
        /// <param name="orderNumberforSO">��ѡ��������ţ�������ṩ�����ʾ���ж���</param>
        /// <param name="statusCodes">��ѡ�����۶���״̬���������״̬ʹ��Ӣ�Ķ��ŷָ���������ṩ�����ʾ����״̬</param>
        /// <returns></returns>
        [HttpGet("pageSize,pageNo,stateCodes,blockNumber,orderNumber,orderNumberforSO")]
        [Route("GetAll")]
        public async Task<IActionResult> GetAll([FromQuery] List<ushort> stateCodes, [FromQuery] string blockNumber, [FromQuery] string orderNumber, [FromQuery] string orderNumberforSO, [FromQuery] int? pageSize = 20, [FromQuery] int? pageNo = 1)
        {
            if (pageSize <= 0 || pageNo <= 0)
                return BadRequest();

            if (!isStatusCodesValid(stateCodes))
                return BadRequest("���ڲ��Ϸ���״̬��");

            IQueryable<WorkOrder> wo = DbContext.WorkOrders.AsQueryable();

            if (stateCodes.Count > 0)
                wo = wo.Where(b => stateCodes.Contains((ushort)b.ManufacturingState));

            blockNumber = blockNumber != null ? blockNumber.Trim() : null;
            if (!string.IsNullOrEmpty(blockNumber))
                wo = wo.Include(w => w.MaterialRequisition)
                    .ThenInclude(mr => mr.Block)
                    .Where(w => w.MaterialRequisition.Block.BlockNumber.Contains(blockNumber));

            orderNumber = orderNumber != null ? orderNumber.Trim() : null;
            if (!string.IsNullOrEmpty(orderNumber))
                wo = wo.Where(w => w.OrderNumber.Contains(orderNumber));

            orderNumberforSO = orderNumberforSO != null ? orderNumberforSO.Trim() : null;
            if (!string.IsNullOrEmpty(orderNumberforSO))
                wo = wo.Include(w => w.SalesOrder)
                        .Where(w => w.SalesOrder.OrderNumber.Contains(orderNumberforSO));

            var orders = wo
                         .Include(w => w.SalesOrderDetail)
                         .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                         .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Category)
                         .OrderByDescending(w => w.CreatedTime)
                         .Skip((pageNo.Value - 1) * pageSize.Value)
                         .Take(pageSize.Value)
                         .AsNoTracking()
                         .ToList();

            foreach (WorkOrder w in orders)
            {
                await Helper.HiddenAmountForSOD(w.SalesOrderDetail, UserManager, User);
            }


            List<WorkOrderForListDTO> woRes = new List<WorkOrderForListDTO>();

            foreach (WorkOrder w in orders)
            {
                WorkOrderForListDTO woDTO = mapper.Map<WorkOrderForListDTO>(w);
                Block block = w.MaterialRequisition != null ? w.MaterialRequisition.Block : null;
                woDTO.BlockNumber = block != null ? block.BlockNumber : null;
                woDTO.BlockCategoryName = block != null ? block.Category.Name : null;
                woRes.Add(woDTO);
            }

            return Success(woRes);
        }

        private bool isStatusCodesValid(List<ushort> statusCodes)
        {
            foreach (ushort code in statusCodes)
            {
                if (!Enum.IsDefined(typeof(ManufacturingState), code))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// ��ȡһ��������������Ϣ
        /// </summary>
        /// <param name="id">Id</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��content�������������������Ϣ������Ҳ�����ӦId���������������ر�׼Error���ݽṹ������ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet("Id")]
        [Route("Get")]
        public async Task<IActionResult> Get(int id)
        {
            if (id <= 0)
                return BadRequest();

            var order = DbContext.WorkOrders
                        .Where(w => w.Id == id)
                        .Include(wo => wo.SalesOrder)
                        .Include(wo => wo.SalesOrderDetail)
                            .ThenInclude(sod => sod.Block)
                        .Include(wo => wo.SalesOrderDetail)
                            .ThenInclude(sod => sod.Bundle)
                        .Include(wo => wo.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.StockingArea)
                        .Include(wo => wo.MaterialRequisition)
                            .ThenInclude(mr => mr.Bundle)
                        .SingleOrDefault();

            if (order == null)
                return Error("�Ҳ���ָ������������");

            await Helper.HiddenAmount(order.SalesOrder, UserManager, User);

            return Success(order);
        }

        /// <summary>
        /// �½���������
        /// </summary>
        /// <param name="workOrder">�����ޱ�JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Create")]
        public async Task<IActionResult> Create([FromBody] WorkOrderCreationModel workOrder)
        {
            workOrder.OrderNumber = workOrder.OrderNumber.Trim();
            workOrder.Priority = workOrder.Priority.Trim();
            var ordNumber = DbContext.WorkOrders
                            .Where(c => c.OrderNumber == workOrder.OrderNumber)
                            .SingleOrDefault();

            if (ordNumber != null)
                return BadRequest("�������Ѵ���");

            if (workOrder.Priority == null)
                return BadRequest("�����빤�����ȼ�");

            var so = DbContext.SalesOrders
                    .Where(s => s.Id == workOrder.SalesOrderId)
                    .Include(d => d.Details)
                    .SingleOrDefault();

            if (so == null)
                return BadRequest("���۶���������");

            var sod = DbContext.SalesOrderDetails
                    .Where(d => d.Id == workOrder.SalesOrderDetailId)
                    .SingleOrDefault();

            if (sod == null)
                return BadRequest("������ϸ������");

            if (sod.OrderId != so.Id)
                return BadRequest("���۶�����������ϸ��ƥ��");

            if (so.OrderType == SalesOrderType.BlockNotInStock || so.OrderType == SalesOrderType.BlockInStock || so.OrderType == SalesOrderType.PolishedSlabInStock)
                return BadRequest("�����п����ϻ��޿����ϻ����п���嶼����Ҫ�½�����");

            if (so.Status != SalesOrderStatus.Approved)
                return BadRequest("�����۶�����״̬�������½���������");

            WorkOrder dbWorkOrder = mapper.Map<WorkOrder>(workOrder);
            dbWorkOrder.Thickness = sod.Specs.Object.Height;
            dbWorkOrder.DeliveryDate = workOrder.DeliveryDate.Value.ToLocalTime();
            await base.InitializeDBRecord(dbWorkOrder);

            DbContext.WorkOrders.Add(dbWorkOrder);
            DbContext.SaveChanges();

            // �������ܺͲֹܷ���Ϣ
            string workOrderInfoPageUrl = "/workOrders/info/" + dbWorkOrder.Id;
            string title = "�µĹ����Ѵ������밲������";
            string text = string.Format("���� {0} ���ύ���밲������", dbWorkOrder.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderInfoPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, workOrder.AuthCode);

            return Success();
        }

        /// <summary>
        /// ������������
        /// </summary>
        /// <param name="workOrder">����������������</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Update")]
        public async Task<IActionResult> Update([FromBody] WorkOrderUpdateModel workOrder)
        {
            workOrder.Priority = workOrder.Priority.Trim();
            var wo = DbContext.WorkOrders
                            .Where(w => w.Id == workOrder.Id)
                            .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (workOrder.Priority == null)
                return BadRequest("�����빤�����ȼ�");

            if (workOrder.DeliveryDate.Value.ToLocalTime() < DateTime.Today)
                return BadRequest("�������ڲ����ǹ�ȥ������");

            if (wo.ManufacturingState == ManufacturingState.Completed || wo.ManufacturingState == ManufacturingState.Cancelled)
                return BadRequest("��������״̬�������޸Ĺ�����Ϣ");

            if (wo.OrderType != workOrder.OrderType)
            {
                if (wo.ManufacturingState == ManufacturingState.NotStarted || wo.ManufacturingState == ManufacturingState.MaterialRequisitionSubmitted
                || wo.ManufacturingState == ManufacturingState.MaterialRequisitioned || wo.ManufacturingState == ManufacturingState.TrimmingDataSubmitted
                || wo.ManufacturingState == ManufacturingState.Trimmed || wo.ManufacturingState == ManufacturingState.SawingDataSubmitted
                || wo.ManufacturingState == ManufacturingState.Sawed)
                    wo.OrderType = workOrder.OrderType;
                else
                    return BadRequest("������������״̬�������޸���������");
            }

            wo.Priority = workOrder.Priority;
            wo.DeliveryDate = workOrder.DeliveryDate.Value.ToLocalTime();
            wo.LastUpdatedTime = DateTime.Now;
            wo.Notes = workOrder.Notes;

            DbContext.SaveChanges();
            // �������ܺͲֹܷ���Ϣ
            string workOrderInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = "���������и���";
            string text = string.Format("���� {0} �и��£���鿴���������ƻ�����Ӧ����", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderInfoPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, workOrder.AuthCode);
            return Success();
        }

        /// <summary>
        /// ��ȡ����������������Ϣ
        /// </summary>
        /// <param name="workOrderId">����Id</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��content���������������Ϣ������Ҳ�����ӦId���������������ر�׼Error���ݽṹ������ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet("Id")]
        [Route("GetMaterialRequisition")]
        public IActionResult GetMaterialRequisition(int workOrderId)
        {
            if (workOrderId <= 0)
                return BadRequest();

            var mr = DbContext.MaterialRequisitions
                        .Where(m => m.WorkOrderId == workOrderId)
                        .Include(m => m.Block)
                            .ThenInclude(b => b.Category)
                        .Include(m => m.Block)
                            .ThenInclude(b => b.StockingArea)
                        .Include(m => m.Block)
                            .ThenInclude(b => b.Grade)
                        .Include(m => m.Bundle)
                        .SingleOrDefault();

            return Success(mr);
        }

        /// <summary>
        /// ������ϵ�
        /// </summary>
        /// <param name="mateReq">���ϵ�Json����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("AddMaterialRequisition")]
        public async Task<IActionResult> AddMaterialRequisition([FromBody] MaterialRequisitionViewModel mateReq)
        {
            if (mateReq.WorkOrderId <= 0)
                return BadRequest("����Id���Ϸ�");

            var wo = DbContext.WorkOrders
                .Where(w => w.Id == mateReq.WorkOrderId)
                .Include(w => w.SalesOrderDetail)
                .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState != ManufacturingState.NotStarted)
                return BadRequest("ֻ����δ��ʼ�����Ĺ���������������ϵ�");

            if (mateReq.BlockId == null && mateReq.BundleId == null)
                return BadRequest("���ϵ�����ָ������Id���ߴ������Id");

            Block block = null;
            if (mateReq.BlockId != null)
            {
                int blockId = mateReq.BlockId.Value;
                block = DbContext.Blocks
                    .Where(b => b.Id == blockId)
                    .SingleOrDefault();

                if (block == null)
                    return Error("���ϵ���ָ���Ļ��ϲ����ڣ�����Id" + blockId);

                var stoneCategory = DbContext.StoneCategories
                    .Include(c => c.BaseCategory)
                    .Where(sc => sc.Id == wo.SalesOrderDetail.Specs.Object.CategoryId)
                    .AsNoTracking()
                    .SingleOrDefault();

                // ȷ����ȡ���ϵ�ʯ��������������ϸ��ָ�����������������Ļ����������࣬�粻���򱨴�
                if (block.CategoryId != stoneCategory.Id && (stoneCategory.BaseCategory == null || stoneCategory.BaseCategory.Id != block.CategoryId))
                    return Error("���ϵ���ָ���Ļ����붩������һ�£�����Id" + blockId);

                if (block.Status != BlockStatus.InStock)
                    return Error("���ϵ���ָ���Ļ��ϲ����ڿ�״̬������Id" + blockId);

                block.Status = BlockStatus.Reserved;
            }

            //todo: ��Ӵ�����д������״̬��鲢�ҽ���״̬���¡�Ŀǰ���������⣺���ϵ�������ȡ������壬������Ҫ���������ϵ��д洢����������ݽṹ

            var dbMR = mapper.Map<MaterialRequisition>(mateReq);
            await InitializeDBRecord(dbMR);

            wo.MaterialRequisition = dbMR;
            wo.ManufacturingState = ManufacturingState.MaterialRequisitionSubmitted;
            DbContext.SaveChanges();

            // �������˷��Ͷ�����Ϣ
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
            {
                string approvalPageUrl = "/workOrders/info/" + dbMR.Id;
                string title = string.Format("�µ����ϵ���Ҫ����\n���ϱ�ţ�{0} ", wo.MaterialRequisition.Block.BlockNumber);
                string text = string.Format("���� {0} �����ϵ����ύ��������", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, approvalPageUrl, RoleDefinition.BlockManager, mateReq.AuthCode);
            }
            else
            {
                string approvalPageUrl = "/workOrders/info/" + dbMR.Id;
                string title = string.Format("�µ����ϵ���Ҫ����\n���ϱ�ţ�{0} ", wo.MaterialRequisition.Block.BlockNumber);
                string text = string.Format("���� {0} �����ϵ����ύ��������", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, approvalPageUrl, RoleDefinition.ProductManager, mateReq.AuthCode);
            }

            return Success();
        }

        /// <summary>
        /// ������������
        /// </summary>
        /// <param name="workOrderId">��������Id</param>
        /// <param name="authCode">���Ͷ�����Ϣ��֤��</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet("workOrderId,authCode")]
        [Route("ApproveMaterialRequisition")]
        public async Task<IActionResult> ApproveMaterialRequisition([FromQuery]int workOrderId, [FromQuery]string authCode)
        {
            if (workOrderId <= 0)
                return BadRequest();

            var wo = DbContext.WorkOrders
                .Where(w => w.Id == workOrderId)
                .Include(w => w.MaterialRequisition)
                    .ThenInclude(mr => mr.Block)
                .Include(w => w.MaterialRequisition)
                    .ThenInclude(mr => mr.Bundle)
                .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState == ManufacturingState.MaterialRequisitioned)
                return Success();

            if (wo.ManufacturingState != ManufacturingState.MaterialRequisitionSubmitted)
                return BadRequest("�ù���������״̬��������׼��������");

            if (wo.MaterialRequisition == null)
                return BadRequest("�ù�����û���ύ���ϵ�");

            if ((wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab) && wo.MaterialRequisition.Block == null)
                return BadRequest("�ù����µ����ϵ�û�л�����Ϣ");

            if (wo.OrderType == WorkOrderType.Tile && wo.MaterialRequisition.Bundle == null)
                return BadRequest("�ù����µ����ϵ�û�д����Ϣ");

            // ���¹���״̬
            wo.ManufacturingState = ManufacturingState.MaterialRequisitioned;

            // ���¶�Ӧ�Ĳ��ϵ�״̬
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
                wo.MaterialRequisition.Block.Status = BlockStatus.Manufacturing;

            if (wo.OrderType == WorkOrderType.Tile)
                wo.MaterialRequisition.Bundle.Status = SlabBundleStatus.CheckedOut;

            Block block = wo.MaterialRequisition.Block;

            block.StockOutOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            block.StockOutTime = DateTime.Now;

            DbContext.SaveChanges();

            // �������˷��Ͷ�����Ϣ
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
            {
                // ��������ܷ��Ͷ�����Ϣ
                string trimmingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("���ϵ�������������ϵ�������\n���ϱ�ţ�{0} ", block.BlockNumber);
                string text = string.Format("���� {0} �����ϵ�������������ϵ�������", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, trimmingInfoPageUrl, RoleDefinition.SawingManager, authCode);
            }
            else
            {
                // �����������ܷ��Ͷ�����Ϣ //todo: ����û�к����������ܵĽ�ɫ
                string trimmingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("���ϵ�������������ϵ�������\n���ϱ�ţ�{0} ", block.BlockNumber);
                string text = string.Format("���� {0} �����ϵ�������������ϵ�������", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, trimmingInfoPageUrl, RoleDefinition.TileQE, authCode);
            }

            return Success();
        }

        /// <summary>
        /// �����ޱ�
        /// </summary>
        /// <param name="trimmingInfo">�����ޱ�JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("UpdateTrimmingData")]
        public async Task<IActionResult> UpdateTrimmingData([FromBody] TrimmingInfo trimmingInfo)
        {
            if (trimmingInfo.WorkOrderId <= 0)
                return BadRequest("������Id���Ϸ�");

            if (trimmingInfo.TrimmingStartTime > trimmingInfo.TrimmingEndTime)
                return BadRequest("�ޱ߿�ʼʱ�䲻�������ޱ߽���ʱ��");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == trimmingInfo.WorkOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState != ManufacturingState.MaterialRequisitioned)
                return BadRequest("�ù�����״̬�������ޱ�");

            if (trimmingInfo.TrimmedHeight <= 0 || trimmingInfo.TrimmedLength <= 0 || trimmingInfo.TrimmedWidth <= 0)
                return BadRequest("�ޱ߳ߴ粻�Ϸ�");

            if (wo.OrderType != WorkOrderType.RawSlab && wo.OrderType != WorkOrderType.PolishedSlab)
                return BadRequest("�������Ͳ�֧���ޱ߲���");

            if (wo.MaterialRequisition == null)
                return BadRequest("�ù���û�ж�Ӧ�����ϵ������ܽ����ޱ߲���");

            if (wo.MaterialRequisition.BlockId == null)
                return BadRequest("�ù��������ϵ���û�ж�Ӧ�Ļ�����Ϣ");

            Block block = wo.MaterialRequisition.Block;

            if (block == null)
                return BadRequest("�ù������ϵ���ָ���Ļ��ϲ�����");

            if (block.Status != BlockStatus.Manufacturing)
                return BadRequest("����״̬�������ޱ�");

            wo.ManufacturingState = ManufacturingState.TrimmingDataSubmitted;
            wo.TrimmingDetails = trimmingInfo.TrimmingDetails;
            wo.TrimmingStartTime = trimmingInfo.TrimmingStartTime.Value.ToLocalTime();
            wo.TrimmingEndTime = trimmingInfo.TrimmingEndTime.Value.ToLocalTime();
            block.TrimmedHeight = trimmingInfo.TrimmedHeight;
            block.TrimmedLength = trimmingInfo.TrimmedLength;
            block.TrimmedWidth = trimmingInfo.TrimmedWidth;
            wo.TrimmingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // ���ޱ��ʼ췢�Ͷ�����Ϣ
            string trimmingQEPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("�ޱ��������ύ��������ޱ��ʼ�\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} ���ޱ��������ύ������ޱߺ�Ļ��Ͻ����ʼ�", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, trimmingQEPageUrl, RoleDefinition.TrimmingQE, trimmingInfo.AuthCode);

            return Success();
        }

        /// <summary>
        /// �����ޱ��ʼ�Աȷ���ޱ�����API
        /// </summary>
        /// <param name="trimmingQE">����JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("TrimmingQE")]
        public async Task<IActionResult> TrimmingQE([FromBody] TrimmingQEViewModel trimmingQE)
        {
            var wo = DbContext.WorkOrders
                     .Where(w => w.Id == trimmingQE.WorkOrderId)
                     .Include(m => m.MaterialRequisition)
                     .SingleOrDefault();

            if (wo == null)
                return BadRequest("ָ��Id����������������");

            if (wo.ManufacturingState != ManufacturingState.TrimmingDataSubmitted)
                return BadRequest("�˻��϶�Ӧ������״̬�������ޱ�");

            if (trimmingQE.TrimmedHeight <= 0 || trimmingQE.TrimmedLength <= 0 || trimmingQE.TrimmedWidth <= 0)
                return BadRequest("�ޱ߳ߴ粻�Ϸ�");

            if (wo.OrderType != WorkOrderType.RawSlab && wo.OrderType != WorkOrderType.PolishedSlab)
                return BadRequest("�������Ͳ�֧���ޱ߲���");

            if (wo.MaterialRequisition == null)
                return BadRequest("�ù���û�ж�Ӧ�����ϵ������ܽ����ޱ߲���");

            if (wo.MaterialRequisition.BlockId == null)
                return BadRequest("�ù��������ϵ���û�ж�Ӧ�Ļ�����Ϣ");

            int blockId = wo.MaterialRequisition.BlockId.Value;

            var block = DbContext.Blocks
                .Where(b => b.Id == blockId)
                .SingleOrDefault();

            if (block == null)
                return BadRequest("�ù������ϵ���ָ���Ļ��ϲ�����");

            if (block.Status != BlockStatus.Manufacturing)
                return BadRequest("����״̬�������ޱ�");

            block.TrimmedHeight = trimmingQE.TrimmedHeight;
            block.TrimmedLength = trimmingQE.TrimmedLength;
            block.TrimmedWidth = trimmingQE.TrimmedWidth;
            wo.BlockOutturnPercentage = Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value) / Helper.CalculateVolume(block.QuarryReportedLength, block.QuarryReportedWidth, block.QuarryReportedHeight);
            wo.BlockOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.BlockOutturnPercentage.Value);
            wo.TrimmingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            wo.ManufacturingState = ManufacturingState.Trimmed;

            DbContext.SaveChanges();

            // ��������ܷ��Ͷ�����Ϣ
            string sawingInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("�ޱ߹�������ɣ�����д��й���\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} ���ޱ߹�������ɣ�����д��й���", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, sawingInfoPageUrl, RoleDefinition.SawingManager, trimmingQE.AuthCode);
            return Success();
        }

        /// <summary>
        /// ���ϴ��
        /// </summary>
        /// <param name="sawingInfo">���ϴ��JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Sawing")]
        public async Task<IActionResult> Sawing([FromBody] SawingInfo sawingInfo)
        {
            if (sawingInfo.WorkOrderId <= 0)
                return BadRequest("������Id���Ϸ�");

            if (sawingInfo.SawingStartTime > sawingInfo.SawingEndTime)
                return BadRequest("��⿪ʼʱ�䲻�����ڴ�����ʱ��");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == sawingInfo.WorkOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState != ManufacturingState.Trimmed)
                return BadRequest("�ù�����״̬��������");

            wo.ManufacturingState = ManufacturingState.SawingDataSubmitted;
            wo.SawingDetails = sawingInfo.SawingDetails;
            wo.SawingStartTime = sawingInfo.SawingStartTime.Value.ToLocalTime();
            wo.SawingEndTime = sawingInfo.SawingEndTime.Value.ToLocalTime();
            wo.SawingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // ������ʼ�Ա���Ͷ�����Ϣ
            string sawingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("����������ύ�������ë���ʼ����\n���ϱ�ţ�{0} ", wo.MaterialRequisition.Block.BlockNumber);
            string text = string.Format("���� {0} �Ĵ���������ύ�������ë���ʼ����", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, sawingQAPageUrl, RoleDefinition.SawingQE, sawingInfo.AuthCode);

            return Success();
        }

        /// <summary> 
        /// �������
        /// </summary>
        /// <param name="inputSB">���JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("SplitBundle")]
        public async Task<IActionResult> SplitBundle([FromBody] StoneBundleSplitingInfo inputSB)
        {
            if (inputSB.TotalSlabCount <= 0)
                return BadRequest("��Ƭ�����Ϸ�");

            if (inputSB.TotalBundleCount <= 0)
                return BadRequest("���������Ϸ�");

            if (inputSB.Thickness <= 0)
                return BadRequest("����Ȳ��Ϸ�");

            var wo = DbContext.WorkOrders
                     .Where(w => w.Id == inputSB.WorkOrderId)
                     .Include(workOrder => workOrder.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                            .ThenInclude(b => b.Bundles)
                     .Include(w => w.SalesOrderDetail)
                     .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.SawingDataSubmitted)
                return BadRequest("�˷�����������״̬���������");

            if (wo.MaterialRequisition == null)
                return BadRequest("�˹��������ϵ�������");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("���ϵ��еĻ��ϲ�����");

            if (wo.SalesOrderDetail == null)
                return BadRequest("����������Ӧ��������ϸ������");

            if (wo.SalesOrderDetail.Specs.Object == null)
                return BadRequest("����������Ӧ��������ϸ��ϸ��Ϣ������");

            if (inputSB.Bundles.Count == 0)
                return BadRequest("�����������ϸ");

            if (inputSB.Bundles.Count != inputSB.TotalBundleCount)
                return BadRequest("����������������������");

            int j = 0;
            int i = 0;

            foreach (BundleInfo b in inputSB.Bundles)
            {
                j++;

                if (b.BundleNo <= 0 || b.BundleNo > inputSB.TotalBundleCount)
                    return BadRequest("���Ų��Ϸ�");

                if (b.BundleNo != j)
                    return BadRequest("�밴˳��༭����");

                if (b.GradeId <= 0)
                    return BadRequest("ʯ������Id���Ϸ�");

                var gra = DbContext.StoneGrades
                          .Where(c => c.Id == b.GradeId)
                          .SingleOrDefault();

                if (gra == null)
                    return BadRequest("ʯ�ĵȼ�������");

                foreach (SlabInfo s in b.Slabs)
                {
                    i++;

                    if (s.SequenceNumber <= 0 || s.SequenceNumber > inputSB.TotalSlabCount)
                        return BadRequest("����Ų��Ϸ�");

                    if (s.SequenceNumber != i)
                        return BadRequest("�밴˳����д����");

                    if (s.Length <= 0 || s.Width <= 0)
                        return BadRequest("���к�ĳߴ粻�Ϸ�");

                    if (s.DeductedLength < 0 || s.DeductedWidth < 0 || s.DeductedLength >= s.Length || s.DeductedWidth >= s.Width)
                        return BadRequest("�۳߳ߴ粻�Ϸ�");

                    if (s.Discarded)
                        s.DiscardedReason = s.DiscardedReason.Trim();

                    if (s.Discarded && string.IsNullOrEmpty(s.DiscardedReason))
                        return BadRequest("�����뱨��ԭ��");
                }
            }

            var blo = wo.MaterialRequisition.Block;

            if (blo == null)
                return BadRequest("�ù������ϵ���ָ���Ļ��ϲ�����");

            if (blo.Bundles.Count > 0)
                return BadRequest("�ù������ϵ���ָ���Ļ������������");

            if (blo.Status != BlockStatus.Manufacturing)
                return BadRequest("����״̬���������");

            blo.TotalSlabNo = inputSB.TotalSlabCount;

            // ��Ʒʯ�����࣬�����ɵĴ����ʹ����ʹ��������ϸ�е�ʯ�����࣬�����ǻ��ϵ�ʯ������
            // ��Ϊ���۶�����ϸ�п��ܻ���˳�кͷ��е�ʯ������
            int productCateId = wo.SalesOrderDetail.Specs.Object.CategoryId;

            foreach (BundleInfo bi in inputSB.Bundles)
            {
                StoneBundle sb = mapper.Map<StoneBundle>(bi);

                await base.InitializeDBRecord(sb);

                sb.Type = SlabType.Raw;
                sb.BlockId = blo.Id;
                sb.CategoryId = productCateId;
                sb.TotalBundleCount = inputSB.TotalBundleCount;
                sb.Thickness = inputSB.Thickness;
                sb.LengthAfterStockIn = blo.TrimmedLength;
                sb.WidthAfterStockIn = blo.TrimmedWidth;
                sb.BlockNumber = blo.BlockNumber;

                int k = 0;
                foreach (Slab slabForDB in sb.Slabs)
                {
                    SlabInfo slabFromClient = bi.Slabs.ToList()[k];
                    slabForDB.Type = SlabType.Raw;
                    slabForDB.BlockId = blo.Id;
                    slabForDB.CategoryId = productCateId;
                    slabForDB.GradeId = sb.GradeId;
                    slabForDB.Thickness = sb.Thickness;
                    slabForDB.LengthAfterSawing = slabFromClient.Length;
                    slabForDB.WidthAfterSawing = slabFromClient.Width;
                    slabForDB.DeductedLength = slabFromClient.DeductedLength;
                    slabForDB.DeductedWidth = slabFromClient.DeductedWidth;
                    slabForDB.Status = slabFromClient.Discarded ? SlabBundleStatus.Discarded : SlabBundleStatus.Manufacturing;
                    slabForDB.DiscardedReason = slabFromClient.Discarded ? slabFromClient.DiscardedReason : null;
                    slabForDB.SawingNote = slabFromClient.SawingNote;
                    await base.InitializeDBRecord(slabForDB);
                    k++;
                }

                sb.TotalSlabCount = Helper.GetAvailableSlabCount(sb);
                sb.Status = Helper.JudgeBundleDiscarded(sb);
                sb.Area = Helper.GetAvailableSlabAreaForBundle(sb);

                foreach (Slab slabForDB in sb.Slabs)
                {
                    slabForDB.ManufacturingState = (slabForDB.Status == SlabBundleStatus.Discarded || !inputSB.SkipFilling) ? ManufacturingState.Sawed : ManufacturingState.Filled;
                }

                sb.ManufacturingState = (sb.Status == SlabBundleStatus.Discarded || !inputSB.SkipFilling) ? ManufacturingState.Sawed : ManufacturingState.Filled;
                blo.Bundles.Add(sb);
            }
            wo.ManufacturingState = inputSB.SkipFilling ? ManufacturingState.Filled : ManufacturingState.Sawed;
            wo.SkipFilling = inputSB.SkipFilling;
            wo.AreaAfterSawing = Helper.GetBlockManufacturingArea(blo, ManufacturingState.Sawed);
            wo.RawSlabOutturnPercentage = wo.AreaAfterSawing.Value / Helper.CalculateVolume(blo.TrimmedLength.Value, blo.TrimmedWidth.Value, blo.TrimmedHeight.Value);
            wo.RawSlabOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.RawSlabOutturnPercentage.Value);
            wo.SawingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();
            // ֱĥ����ĥ���ʼ췢�Ͷ�����Ϣ
            if (inputSB.SkipFilling)
            {
                string PolishingingQAPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("���й�������ɣ��˿Ż��ϵĴ�����貹���������ĥ�׹���\n���ϱ�ţ�{0} ", blo.BlockNumber);
                string text = string.Format("���� {0} �Ĵ��й�������ɣ��˿Ż��ϵĴ�����貹���������ĥ�׹���", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, PolishingingQAPageUrl, RoleDefinition.PolishingQE, inputSB.AuthCode);

            }
            // ����в��������������ܷ��Ͷ�����Ϣ
            else
            {
                string fillingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("���й�������ɣ��˿Ż��ϵĴ�����貹���������ĥ�׹���\n���ϱ�ţ�{0} ", blo.BlockNumber);
                string text = string.Format("���� {0} �Ĵ��й�������ɣ�����в���", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, fillingInfoPageUrl, RoleDefinition.FillingManager, inputSB.AuthCode);

            }

            return Success();
        }

        /// <summary>
        /// ��岹��
        /// </summary>
        /// <param name="fillingInfo">��岹��JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Filling")]
        public async Task<IActionResult> Filling([FromBody] FillingInfo fillingInfo)
        {
            if (fillingInfo.WorkOrderId <= 0)
                return BadRequest("������Id���Ϸ�");

            if (fillingInfo.FillingStartTime > fillingInfo.FillingEndTime)
                return BadRequest("������ʼʱ�䲻�����ڲ�������ʱ��");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == fillingInfo.WorkOrderId)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState != ManufacturingState.Sawed && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�ù�����״̬��������");

            if (wo.MaterialRequisition == null)
                return BadRequest("���������������ϵ�������");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("�����ϵ��Ļ��ϲ�����");

            Block block = wo.MaterialRequisition.Block;

            foreach (StoneBundle bi in block.Bundles)
            {
                if (bi.Status != SlabBundleStatus.Manufacturing)
                    continue;

                if (bi.ManufacturingState != ManufacturingState.Sawed && bi.ManufacturingState != ManufacturingState.Filled)
                    continue;

                if (bi.ManufacturingState == ManufacturingState.Sawed)
                    bi.ManufacturingState = ManufacturingState.FillingDataSubmitted;

                foreach (Slab slabForDB in bi.Slabs)
                {
                    if (slabForDB.Status == SlabBundleStatus.Manufacturing && slabForDB.ManufacturingState == ManufacturingState.Sawed)
                        slabForDB.ManufacturingState = ManufacturingState.FillingDataSubmitted;
                }
            }

            if (wo.ManufacturingState == ManufacturingState.Sawed)
                wo.ManufacturingState = ManufacturingState.FillingDataSubmitted;
            wo.FillingDetails = fillingInfo.FillingDetails;
            wo.FillingStartTime = fillingInfo.FillingStartTime.Value.ToLocalTime();
            wo.FillingEndTime = fillingInfo.FillingEndTime.Value.ToLocalTime();
            wo.FillingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            wo.LastUpdatedTime = DateTime.Now;

            DbContext.SaveChanges();

            // �������ʼ췢�Ͷ�����Ϣ
            string fillingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("�����������ύ������в����ʼ�\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} �Ĳ����������ύ������в����ʼ�", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, fillingQAPageUrl, RoleDefinition.FillingQE, fillingInfo.AuthCode);

            return Success();
        }

        /// <summary>
        /// ���ĥ��
        /// </summary>
        /// <param name="polishingInfo">�����ޱ�JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Polishing")]
        public async Task<IActionResult> Polishing([FromBody] PolishingInfo polishingInfo)
        {
            if (polishingInfo.WorkOrderId <= 0)
                return BadRequest("������Id���Ϸ�");

            if (polishingInfo.PolishingStartTime > polishingInfo.PolishingEndTime)
                return BadRequest("ĥ�׿�ʼʱ�䲻������ĥ�׽���ʱ��");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == polishingInfo.WorkOrderId)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("����������");

            if (wo.ManufacturingState != ManufacturingState.PolishingQEFinished)
                return BadRequest("�ù�����״̬������ĥ��");

            if (wo.MaterialRequisition == null)
                return BadRequest("���������������ϵ�������");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("�����ϵ��Ļ��ϲ�����");

            Block block = wo.MaterialRequisition.Block;

            foreach (StoneBundle bi in block.Bundles)
            {
                if (bi.Status == SlabBundleStatus.Discarded)
                    continue;

                if (bi.ManufacturingState == ManufacturingState.Completed)
                {
                    foreach (Slab slabForDB in bi.Slabs)
                    {
                        if (slabForDB.ManufacturingState != ManufacturingState.Completed && slabForDB.Status != SlabBundleStatus.Discarded)
                            return BadRequest("�˴��״̬��������ɹ���");
                    }
                }
                else
                    return BadRequest("�������״̬��������ɹ���");
            }

            block.Status = BlockStatus.Processed;
            wo.ManufacturingState = ManufacturingState.Completed;
            wo.PolishingDetails = polishingInfo.PolishingDetails;
            wo.PolishingStartTime = polishingInfo.PolishingStartTime.Value.ToLocalTime();
            wo.PolishingEndTime = polishingInfo.PolishingEndTime.Value.ToLocalTime();
            wo.PolishingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // ����Ʒ��ܷ��Ͷ�����Ϣ
            string productStockingInPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("�����������ɣ�����й�����\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} �Ĺ����������ɣ�����й�����", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, productStockingInPageUrl, new string[] { RoleDefinition.ProductManager, RoleDefinition.PackagingManger, RoleDefinition.FactoryManager }, polishingInfo.AuthCode);
            return Success();
        }

        /// <summary>
        /// �����ʼ�
        /// </summary>
        /// <param name="fillingQE">���JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("FillingQE")]
        public async Task<IActionResult> FillingQE([FromBody] SlabPolishingOrFillingQEViewModel fillingQE)
        {
            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == fillingQE.WorkOrderId)
                    .Include(w => w.SalesOrder)
                    .Include(w => w.SalesOrderDetail)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(m => m.Block)
                            .ThenInclude(b => b.Bundles)
                                .ThenInclude(bd => bd.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.FillingDataSubmitted && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹�����״̬��������");

            if (fillingQE.Length <= 0 || fillingQE.Width <= 0)
                return BadRequest("�����ߴ粻�Ϸ�");

            if (fillingQE.DeductedLength < 0 || fillingQE.DeductedWidth < 0 || fillingQE.DeductedLength >= fillingQE.Length || fillingQE.DeductedWidth >= fillingQE.Width)
                return BadRequest("�۳߳ߴ粻�Ϸ�");

            var so = wo.SalesOrder;

            if (so == null)
                return BadRequest("���۶���������");

            if (so.OrderType != SalesOrderType.PolishedSlabNotInStock)
                return BadRequest("�����۶����������貹������");

            var sod = wo.SalesOrderDetail;

            if (sod == null)
                return BadRequest("���۶�����ϸ������");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("�˹�����Ӧ�����ϵ�������");

            var block = mr.Block;

            if (block == null)
                return BadRequest("���ϵ��еĻ��ϲ�����");

            if (fillingQE.Discarded)
                fillingQE.DiscardedReason = fillingQE.DiscardedReason.Trim();

            if (fillingQE.Discarded && string.IsNullOrEmpty(fillingQE.DiscardedReason))
                return BadRequest("�����뱨��ԭ��");

            var slab = DbContext.Slabs
                      .Where(s => s.Id == fillingQE.SlabId)
                      .Include(s => s.Bundle)
                        .ThenInclude(b => b.Slabs)
                      .SingleOrDefault();

            if (slab == null)
                return BadRequest("��岻����");

            if (slab.Bundle == null)
                return BadRequest("���Ų�����");

            var stoneBundles = block.Bundles;

            if (stoneBundles.Count <= 0)
                return BadRequest("���ϵ��еĻ���δ���������");

            if (!(stoneBundles.Contains(slab.Bundle)))
                return BadRequest("��岻�ڴ˷������������������Ĵ����");

            if (slab.Status != SlabBundleStatus.Discarded && slab.Status != SlabBundleStatus.Manufacturing)
                return BadRequest("���״̬��������");

            if (slab.ManufacturingState != ManufacturingState.FillingDataSubmitted && slab.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�������״̬��������");

            slab.Status = fillingQE.Discarded ? SlabBundleStatus.Discarded : SlabBundleStatus.Manufacturing;
            slab.DiscardedReason = fillingQE.Discarded ? fillingQE.DiscardedReason : null;
            slab.ManufacturingState = fillingQE.Discarded ? ManufacturingState.FillingDataSubmitted : ManufacturingState.Filled;

            slab.LengthAfterFilling = fillingQE.Length;//�������������ύ��Ƭ���ݣ��շ��ĳߴ��賤��5cm,���3cm
            slab.WidthAfterFilling = fillingQE.Width;
            slab.DeductedLength = fillingQE.DeductedLength;
            slab.DeductedWidth = fillingQE.DeductedWidth;
            slab.FillingNote = fillingQE.FillingNote;

            slab.Bundle.Status = Helper.JudgeBundleDiscarded(slab.Bundle);
            slab.Bundle.Area = Helper.GetAvailableSlabAreaForBundle(slab.Bundle);
            slab.Bundle.TotalSlabCount = Helper.GetAvailableSlabCount(slab.Bundle);
            wo.AreaAfterFilling = Helper.GetBlockManufacturingArea(block, ManufacturingState.Filled);

            wo.FillingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();
            return Success();
        }

        /// <summary>
        /// �����ʼ�ȷ�Ϲ������
        /// </summary>
        /// <param name="workOrderId">��������IdJSON����</param>
        /// <param name="authCode">���Ͷ�����Ϣ��֤��</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet("workOrderId,authCode")]
        [Route("FillingOver")]
        public async Task<IActionResult> FillingOver([FromQuery] int workOrderId, [FromQuery] String authCode)
        {
            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == workOrderId)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.FillingDataSubmitted && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹�����״̬��������");

            if (wo.MaterialRequisition == null)
                return BadRequest("���������������ϵ�������");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("�����ϵ��Ļ��ϲ�����");

            Block block = wo.MaterialRequisition.Block;

            foreach (StoneBundle bi in block.Bundles)
            {
                if (bi.Status != SlabBundleStatus.Manufacturing)
                    continue;

                if (bi.ManufacturingState != ManufacturingState.FillingDataSubmitted && bi.ManufacturingState != ManufacturingState.Filled)
                    continue;

                if (bi.ManufacturingState == ManufacturingState.FillingDataSubmitted)
                    bi.ManufacturingState = ManufacturingState.Filled;

                foreach (Slab slabForDB in bi.Slabs)
                {
                    if (slabForDB.Status == SlabBundleStatus.Manufacturing && slabForDB.ManufacturingState == ManufacturingState.FillingDataSubmitted)
                    {
                        // ��������󳤶ȺͿ��Ϊ�գ���ʾ�ÿ���û�б������ʼ쵥Ƭ�ύ�����ݣ����Բ�������ʱҪ�����е�״̬Ϊ�����������ύ�Ĵ���������ߴ�ĳ�ʼ��
                        // �������ߴ��ʼ��Ϊ���к�ĳߴ磬�����в������ݵĴ�������κβ����ߴ�ĸ���
                        if (slabForDB.LengthAfterFilling == null)
                            slabForDB.LengthAfterFilling = slabForDB.LengthAfterSawing;
                        if (slabForDB.WidthAfterFilling == null)
                            slabForDB.WidthAfterFilling = slabForDB.WidthAfterSawing;
                        slabForDB.ManufacturingState = ManufacturingState.Filled;
                    }
                }
            }

            if (wo.ManufacturingState == ManufacturingState.FillingDataSubmitted)
                wo.ManufacturingState = ManufacturingState.Filled;

            wo.FillingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            wo.AreaAfterFilling = Helper.GetBlockManufacturingArea(block, ManufacturingState.Filled);
            wo.LastUpdatedTime = DateTime.Now;

            DbContext.SaveChanges();
            // ��ĥ���ʼ췢�Ͷ�����Ϣ
            string polishingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("������������ɣ������ĥ�׹���\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} �Ĳ�����������ɣ������ĥ�׹���", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, polishingQAPageUrl, RoleDefinition.PolishingQE, authCode);

            return Success();
        }

        /// <summary>
        /// ĥ���ʼ�
        /// </summary>
        /// <param name="polishingQE">���JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("PolishingQE")]
        public async Task<IActionResult> PolishingQE([FromBody] SlabPolishingOrFillingQEViewModel polishingQE)
        {
            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == polishingQE.WorkOrderId)
                    .Include(w => w.SalesOrder)
                    .Include(w => w.SalesOrderDetail)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(m => m.Block)
                            .ThenInclude(b => b.Bundles)
                                .ThenInclude(bd => bd.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹�����״̬������ĥ��");

            if (polishingQE.Length <= 0 || polishingQE.Width <= 0)
                return BadRequest("ĥ�׳ߴ粻�Ϸ�");

            if (polishingQE.DeductedLength < 0 || polishingQE.DeductedWidth < 0 || polishingQE.DeductedLength >= polishingQE.Length || polishingQE.DeductedWidth >= polishingQE.Width)
                return BadRequest("�۳߳ߴ粻�Ϸ�");

            var so = wo.SalesOrder;

            if (so == null)
                return BadRequest("���۶���������");

            if (so.OrderType != SalesOrderType.PolishedSlabNotInStock)
                return BadRequest("�����۶�����������ĥ�׹���");

            var sod = wo.SalesOrderDetail;

            if (sod == null)
                return BadRequest("���۶�����ϸ������");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("�˹�����Ӧ�����ϵ�������");

            var block = mr.Block;

            if (block == null)
                return BadRequest("���ϵ��еĻ��ϲ�����");

            if (polishingQE.Discarded)
                polishingQE.DiscardedReason = polishingQE.DiscardedReason.Trim();

            if (polishingQE.Discarded && string.IsNullOrEmpty(polishingQE.DiscardedReason))
                return BadRequest("�����뱨��ԭ��");

            var slab = DbContext.Slabs
                      .Where(s => s.Id == polishingQE.SlabId)
                      .Include(s => s.Bundle)
                        .ThenInclude(b => b.Slabs)
                      .SingleOrDefault();

            if (slab == null)
                return BadRequest("��岻����");

            if (slab.Bundle == null)
                return BadRequest("���Ų�����");

            var stoneBundles = block.Bundles;

            if (stoneBundles.Count <= 0)
                return BadRequest("���ϵ��еĻ���δ���������");

            if (!(stoneBundles.Contains(slab.Bundle)))
                return BadRequest("��岻�ڴ˷������������������Ĵ����");

            if (slab.Status != SlabBundleStatus.Manufacturing && slab.Status != SlabBundleStatus.Discarded)
                return BadRequest("���״̬������ĥ��");

            if (slab.ManufacturingState != ManufacturingState.Filled && slab.ManufacturingState != ManufacturingState.Completed)
                return BadRequest("�������״̬������ĥ��");

            slab.Status = polishingQE.Discarded ? SlabBundleStatus.Discarded : SlabBundleStatus.Manufacturing;
            slab.DiscardedReason = polishingQE.Discarded ? polishingQE.DiscardedReason : null;
            slab.ManufacturingState = polishingQE.Discarded ? ManufacturingState.Filled : ManufacturingState.Completed;

            slab.LengthAfterPolishing = polishingQE.Length;
            slab.WidthAfterPolishing = polishingQE.Width;
            slab.DeductedLength = polishingQE.DeductedLength;
            slab.DeductedWidth = polishingQE.DeductedWidth;
            slab.Type = SlabType.Polished;
            slab.PolishingNote = polishingQE.PolishingNote;

            slab.Bundle.Status = Helper.JudgeBundleDiscarded(slab.Bundle);
            slab.Bundle.Area = Helper.GetAvailableSlabAreaForBundle(slab.Bundle);
            slab.Bundle.TotalSlabCount = Helper.GetAvailableSlabCount(slab.Bundle);
            slab.LastUpdatedTime = DateTime.Now;
            wo.AreaAfterPolishing = Helper.GetBlockManufacturingArea(block, ManufacturingState.Completed);

            wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            wo.LastUpdatedTime = DateTime.Now;

            DbContext.SaveChanges();
            return Success();
        }

        /// <summary>
        /// ĥ�׺��ÿ����嶨�ȼ�
        /// </summary>
        /// <param name="bundleGradeQE">���JSON����</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("BundleGradeQE")]
        public async Task<IActionResult> BundleGradeQE([FromBody] BundleGradePolishingQE bundleGradeQE)
        {
            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == bundleGradeQE.WorkOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                            .ThenInclude(b => b.Bundles)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹�����״̬������ĥ��");

            var bundle = DbContext.StoneBundles
                         .Where(b => b.Id == bundleGradeQE.BundleId)
                         .Include(b => b.Slabs)
                         .SingleOrDefault();

            if (bundle == null)
                return BadRequest("������岻����");

            if (bundle.Slabs.Count <= 0)
                return BadRequest("������û�д��");

            if (bundle.Status != SlabBundleStatus.Manufacturing)
                return BadRequest("�������״̬�������ȼ�");

            var gra = DbContext.StoneGrades
                    .Where(g => g.Id == bundleGradeQE.GradeId)
                    .SingleOrDefault();

            if (gra == null)
                return BadRequest("ʯ�ĵȼ�������");


            bundle.GradeId = bundleGradeQE.GradeId;

            if (bundle.Status == SlabBundleStatus.Manufacturing && (bundle.ManufacturingState == ManufacturingState.Filled || bundle.ManufacturingState == ManufacturingState.Completed))
            {
                foreach (Slab slabForDB in bundle.Slabs)
                {
                    if (slabForDB.Status == SlabBundleStatus.Discarded)
                        continue;

                    if (slabForDB.ManufacturingState != ManufacturingState.Completed && slabForDB.ManufacturingState != ManufacturingState.Filled)
                        return BadRequest("�˴���״̬���������������״̬");

                    slabForDB.ManufacturingState = ManufacturingState.Completed;
                    slabForDB.GradeId = bundle.GradeId;
                }
            }
            else
                return BadRequest("�������״̬���������������״̬");

            bundle.ManufacturingState = ManufacturingState.Completed;
            bundle.Type = SlabType.Polished;
            bundle.LastUpdatedTime = DateTime.Now;
            wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();
            return Success();
        }

        /// <summary>
        /// ĥ���ʼ�ȷ�Ϲ������
        /// </summary>
        /// <param name="workOrderId">��������IdJSON����</param>
        /// <param name="authCode">���Ͷ�����Ϣ��֤��</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet("workOrderId")]
        [Route("PolishingOver")]
        public async Task<IActionResult> PolishingOver([FromQuery] int workOrderId, [FromQuery] String authCode)
        {
            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == workOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹�����״̬������ĥ��");

            Block block = wo.MaterialRequisition.Block;

            wo.ManufacturingState = ManufacturingState.PolishingQEFinished;
            wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            wo.PolishedSlabOutturnPercentage = wo.AreaAfterPolishing.Value / Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value);
            wo.PolishedSlabOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.PolishedSlabOutturnPercentage.Value);

            DbContext.SaveChanges();

            // ��ĥ�����ܷ��Ͷ�����Ϣ
            string polishingInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("ĥ���ʼ��������ύ�������ĥ��ȷ��\n���ϱ�ţ�{0} ", block.BlockNumber);
            string text = string.Format("���� {0} ��ĥ���ʼ��������ύ�������ĥ��ȷ��", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, polishingInfoPageUrl, RoleDefinition.SlabPolishingManager, authCode);
            return Success();
        }

        /// <summary>
        /// ȡ����������
        /// </summary>
        /// <param name="cancel">���۶���JSON���ݣ��������б����ṩId</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Cancel")]
        public async Task<IActionResult> Cancel([FromBody] WorkOrderCancelModel cancel)
        {
            if (cancel.CancelReason == null)
                return BadRequest("������ȡ��������ԭ��");

            var wo = DbContext.WorkOrders
                     .Where(s => s.Id == cancel.WorkOrderId)
                     .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("��������������");

            if ((wo.ManufacturingState != ManufacturingState.NotStarted && wo.ManufacturingState != ManufacturingState.MaterialRequisitionSubmitted
             && wo.ManufacturingState != ManufacturingState.MaterialRequisitioned) || wo.ManufacturingState == ManufacturingState.Cancelled)
                return BadRequest("����������״̬������ȡ����������");

            if (wo.ManufacturingState != ManufacturingState.NotStarted)
            {
                Block block = wo.MaterialRequisition.Block;
                if (block.Status == BlockStatus.Reserved || block.Status == BlockStatus.Manufacturing)
                    block.Status = BlockStatus.InStock;

            }

            wo.ManufacturingState = ManufacturingState.Cancelled;
            wo.CancelReason = cancel.CancelReason;
            wo.LastUpdatedTime = DateTime.Now;

            DbContext.SaveChanges();

            // �������Ա���Ͷ�����Ϣ
            string workOrderCancelPageUrl = "/workOrders/info/" + wo.Id;
            string title = "";
            if (wo.MaterialRequisition != null)
            {
                if (wo.MaterialRequisition.Block != null)
                    title = string.Format("������ȡ��\n���ϱ�ţ�{0} ", wo.MaterialRequisition.Block.BlockNumber);
            }
            else
                title = "������ȡ��";
            string text = string.Format("���� {0} ��ȡ������鿴��ֹͣ����", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderCancelPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, cancel.AuthCode);

            return Success();

        }

        /// <summary>
        /// ��ȡ�ҵ�����
        /// </summary>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet()]
        [Route("GetMyWorkOrders")]
        public async Task<IActionResult> GetMyWorkOrders()
        {
            AnDaUser user = await UserManager.FindByNameAsync(User.Identity.Name);

            IList<string> roleNames = await UserManager.GetRolesAsync(user);

            List<ManufacturingState> manufacturingStates = new List<ManufacturingState>();

            manufacturingStates = getRoleAssignedManufacturingState(roleNames);

            var workOrderList = DbContext.WorkOrders
                        .Where(w => manufacturingStates.Contains(w.ManufacturingState))
                        .Include(w => w.SalesOrderDetail)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Category)
                        .AsNoTracking()
                        .ToList();

            if (roleNames.Contains(RoleDefinition.FillingManager) || roleNames.Contains(RoleDefinition.FillingQE))
            {
                var extraWOList = DbContext.WorkOrders
                    .Where(w => w.ManufacturingState == ManufacturingState.Filled)
                    .Include(w => w.SalesOrderDetail)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Category)
                        .AsNoTracking()
                        .ToList();
                var ms = roleNames.Contains(RoleDefinition.FillingManager) ? ManufacturingState.Sawed : ManufacturingState.FillingDataSubmitted;

                extraWOList = extraWOList.FindAll(wo =>
                {
                    if (wo.MaterialRequisition == null)
                        return false;
                    Block block = wo.MaterialRequisition.Block;
                    if (block == null || block.Bundles.Count == 0)
                        return false;

                    return block.Bundles.Any(sb =>
                    {
                        return sb.Slabs.Any(s => { return s.ManufacturingState == ms && s.Status == SlabBundleStatus.Manufacturing; });
                    });
                });

                extraWOList.ForEach(ewo =>
                {
                    if (workOrderList.Find(wo => wo.Id == ewo.Id) == null)
                        workOrderList.Add(ewo);
                });
            }

            workOrderList = workOrderList.OrderByDescending(s => s.CreatedTime).ToList();

            List<WorkOrderForListDTO> woRes = new List<WorkOrderForListDTO>();

            foreach (WorkOrder wo in workOrderList)
            {
                WorkOrderForListDTO woDTO = mapper.Map<WorkOrderForListDTO>(wo);
                Block block = wo.MaterialRequisition != null ? wo.MaterialRequisition.Block : null;
                woDTO.BlockNumber = block != null ? block.BlockNumber : null;
                woDTO.BlockCategoryName = block != null ? block.Category.Name : null;
                woRes.Add(woDTO);
            }

            return Success(woRes);

        }

        /// <summary>
        /// ��ȡ�����������
        /// </summary>
        /// <returns>����ɹ����򷵻���ţ�ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet()]
        [Route("GetWorkOrderNumber")]
        public IActionResult GetWorkOrderNumber()
        {
            string woFormat = "ADWO{0}-{1}";
            string woPrefix = "workOrder";
            var wn = DbContext.SerialNumbers
                     .Where(s => s.SNName == woPrefix)
                     .SingleOrDefault();
            if (wn == null)
            {
                wn = new SerialNumber() { SNName = woPrefix, Number = 0 };
                DbContext.SerialNumbers.Add(wn);
            }

            wn.Number = wn.Number + 1;
            DbContext.SaveChanges();

            string time = DateTime.Now.ToString("yyyyMMdd");
            string WN = wn.Number.ToString().PadLeft(2, '0');

            string woSN = string.Format(woFormat, time, WN);

            return Success(woSN);
        }

        /// <summary>
        /// �������ز���
        /// </summary>
        /// <param name="bundleId">�������Id</param>
        /// <param name="authCode">���Ͷ�����Ϣ��֤��</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet()]
        [Route("BundleGoBackFilling")]
        public async Task<IActionResult> BundleGoBackFilling([FromQuery] int bundleId, [FromQuery] string authCode)
        {
            if (bundleId <= 0)
                return BadRequest("�������Id������");

            var sb = DbContext.StoneBundles
                            .Where(b => b.Id == bundleId)
                            .Include(b => b.Block)
                            .Include(b => b.Slabs)
                            .SingleOrDefault();

            if (sb == null)
                return BadRequest("������岻����");

            if (sb.TotalSlabCount == 0)
                return BadRequest("�������û�п��õĴ��");

            var mr = DbContext.MaterialRequisitions
                    .Where(m => m.BlockId == sb.BlockId)
                    .Include(m => m.WorkOrder)
                    .SingleOrDefault();

            if (mr == null)
                return BadRequest("���ϵ�������");

            if (mr.WorkOrder == null)
                return BadRequest("��������������");

            if (mr.WorkOrder.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹���״̬�������ز���");

            if (sb.Status != SlabBundleStatus.Manufacturing && sb.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˴���״̬�������ز���");

            sb.ManufacturingState = ManufacturingState.Sawed;
            foreach (Slab slab in sb.Slabs)
            {
                if (slab.Status == SlabBundleStatus.Manufacturing && slab.ManufacturingState == ManufacturingState.Filled)
                    slab.ManufacturingState = ManufacturingState.Sawed;
            }
            DbContext.SaveChanges();

            // ���������ܷ���Ϣ
            string workOrderPageUrl = "/workorders/info/" + mr.WorkOrder.Id;
            string title = "��巵�ز���";
            string text = string.Format("{0} ������ĥ���ʼ촦���ز���������ɲ�������²�������", string.Format("{0} {1}-{2}", sb.BlockNumber, sb.TotalBundleCount, sb.BundleNo));
            await SendDingtalkLinkMessage(title, text, workOrderPageUrl, RoleDefinition.FillingManager, authCode);
            return Success();
        }

        /// <summary>
        /// ��Ƭ���ز���
        /// </summary>
        /// <param name="slabId">���Id</param>
        /// <param name="authCode">������ʱ��Ȩ�룬����������򲻽��ж���֪ͨ</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet()]
        [Route("SlabGoBackFilling")]
        public async Task<IActionResult> SlabGoBackFilling([FromQuery] int slabId, [FromQuery] string authCode)
        {
            if (slabId <= 0)
                return BadRequest("�������Id������");

            var slab = DbContext.Slabs
                       .Where(s => s.Id == slabId)
                       .Include(s => s.Bundle)
                       .Include(s => s.Block)
                       .SingleOrDefault();

            if (slab == null)
                return BadRequest("�˴�岻����");

            var bundle = DbContext.StoneBundles
                        .Where(b => b.Id == slab.BundleId)
                        .Include(b => b.Slabs)
                        .SingleOrDefault();

            if (bundle == null)
                return BadRequest("�˴���Ӧ����������");

            if (bundle.Status != SlabBundleStatus.Manufacturing && bundle.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("����Ӧ����״̬�������ز���");

            if (slab.Status != SlabBundleStatus.Manufacturing && slab.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˴���״̬�������ز���");

            var mr = DbContext.MaterialRequisitions
                    .Where(m => m.BlockId == slab.BlockId)
                    .Include(m => m.WorkOrder)
                    .SingleOrDefault();

            if (mr == null)
                return BadRequest("���ϵ�������");

            if (mr.WorkOrder == null)
                return BadRequest("��������������");

            if (mr.WorkOrder.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("�˹���״̬�������ز���");

            slab.ManufacturingState = ManufacturingState.Sawed;
            bundle.ManufacturingState = Helper.JudgeBundleGoBackFilling(bundle);

            DbContext.SaveChanges();

            // ���������ܷ���Ϣ
            string workOrderPageUrl = "/workorders/info/" + mr.WorkOrder.Id;
            string title = "��巵�ز���";
            string text = string.Format("{0} ���е����Ϊ {1} �Ĵ���ĥ���ʼ촦���ز���������ɲ�������²�������", string.Format("{0} {1}-{2}", bundle.BlockNumber, bundle.TotalBundleCount, bundle.BundleNo), slab.SequenceNumber);
            await SendDingtalkLinkMessage(title, text, workOrderPageUrl, RoleDefinition.FillingManager, authCode);

            return Success();
        }

        /// <summary>
        /// ����ָ���Ĺ����еĻ���
        /// </summary>
        /// <param name="workOrderId">����JSON���ݣ��������б����ṩId</param>
        /// <param name="discardedReason">����ԭ���������б����ṩ����ԭ��</param>
        /// <param name="authCode">���Ͷ�����Ϣ��֤��</param>
        /// <returns>����ɹ����򷵻ر�׼Success���ݽṹ��ʧ�ܷ�����Ӧ��HTTP�������</returns>
        [HttpGet]
        [Route("DiscardBlock")]
        public async Task<IActionResult> DiscardBlock([FromQuery] int workOrderId, [FromQuery] string discardedReason, [FromQuery] string authCode)
        {
            if (workOrderId <= 0)
                return BadRequest("����Id���Ϸ�");

            discardedReason = discardedReason.Trim();
            if (string.IsNullOrEmpty(discardedReason))
                return BadRequest("���ϱ���ԭ����Ϊ��");

            var wo = DbContext.WorkOrders
                .Where(w => w.Id == workOrderId)
                .Include(w => w.MaterialRequisition)
                    .ThenInclude(m => m.Block)
                .SingleOrDefault();

            if (wo == null)
                return BadRequest("����Id��Ӧ�Ĺ���������");

            if (!(wo.ManufacturingState == ManufacturingState.TrimmingDataSubmitted || wo.ManufacturingState == ManufacturingState.SawingDataSubmitted))
                return BadRequest("��������������״̬���ܽ��л��ϱ��ϲ���");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("����û�ж�Ӧ�����ϵ�");

            if (mr.Block == null)
                return BadRequest("������Ӧ�����ϵ����ǻ������ϵ��������ϵ��еĻ��ϲ�����");

            var blo = mr.Block;

            if (blo.Status != BlockStatus.Manufacturing)
                return BadRequest("�˻��ϲ�������״̬�����ܽ��б��ϲ���");

            ManufacturingState status = wo.ManufacturingState;
            wo.ManufacturingState = ManufacturingState.Completed;
            wo.BlockDiscarded = true;
            wo.LastUpdatedTime = blo.LastUpdatedTime = DateTime.Now;
            blo.Status = BlockStatus.Discarded;
            blo.DiscardedReason = discardedReason;
            DbContext.SaveChanges();

            // �������Ա���Ͷ�����Ϣ
            string processName = (status == ManufacturingState.TrimmingDataSubmitted) ? "�ޱ߹���" : "���й���";

            string workOrderCancelPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("���ϱ���\n���ϱ�ţ�{0} ", blo.BlockNumber);
            string text = string.Format("���� {0} �Ļ��� {1} ��{2}���ϣ���鿴��ȡ�������ĺ��������ƻ�", wo.OrderNumber, blo.BlockNumber, processName);
            await SendDingtalkLinkMessage(title, text, workOrderCancelPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, authCode);

            return Success();
        }

        /// <summary>
        /// ������������
        /// </summary>
        /// <param name="specFile">����������ļ���Ϣ</param>
        /// <returns>�����������а�������ʹ�����Ϣ</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("ImportBundleInStock")]
        [Authorize(Roles = RoleDefinition.DataOperator)]
        public async Task<IActionResult> ImportBundleInStock([FromBody] BundleSpecFileViewModel specFile)
        {
            string base64Content = specFile.FileContent;

            if (string.IsNullOrEmpty(specFile.FileName))
                return BadRequest("�ļ�������Ϊ��");

            if (!specFile.FileName.EndsWith(".xlsx"))
                return Error("�ϴ����ļ�������Excel 2007�Ժ��ʽ��.xlsx�ļ�");

            if (string.IsNullOrEmpty(base64Content))
                return BadRequest("�ļ����ݲ���Ϊ��");

            byte[] fileBytes = null;
            try
            {
                fileBytes = Convert.FromBase64String(base64Content);
            }
            catch
            {
                return Error("�ļ����ݲ��Ϸ�������ʧ��");
            }

            List<string> dataParsingMsgs = new List<string>();
            List<string> dataUpdatingMsgs = new List<string>();
            List<StoneBundle> importBundles = new List<StoneBundle>();

            logger.LogTrace("���ļ��ж�ȡ������Ϣ");

            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                PolishedBundleSpecSheetProcessingResult result = bundleInStockSpecSheetSVC.GetBundleInStockSpecs(ms);
                dataParsingMsgs = result.ErrorMessages;
                importBundles = result.Bundles;
            }

            string dataOperatorName = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            foreach (StoneBundle importBundle in importBundles)
            {
                string bundleNo = string.Format("{0} #{1}-{2}", importBundle.BlockNumber, importBundle.TotalBundleCount, importBundle.BundleNo);
                logger.LogTrace("������������� - {bundleNo}", bundleNo);

                // ������Ӧ�����ݿ���Ϣ
                var bundleInDB = DbContext.StoneBundles
                    .Include(sb => sb.Slabs)
                    .Include(sb => sb.Category)
                    .Include(sb => sb.Grade)
                    .Where(sb => sb.BlockNumber == importBundle.BlockNumber && sb.BundleNo == importBundle.BundleNo)
                    .SingleOrDefault();

                // ���ݿ��в�����������壬ֱ�ӵ���
                if (bundleInDB == null)
                {
                    // �Ҷ�Ӧ�Ļ��ϣ�����ҵ���ѻ��ϵ�״̬����Ϊ���������
                    var block = DbContext.Blocks
                        .Where(b => b.BlockNumber == importBundle.BlockNumber)
                        .SingleOrDefault();
                    if (block != null && block.Status != BlockStatus.InStock && block.Status != BlockStatus.Processed)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("������Ӧ�Ļ�����ϵͳ���Ѿ����ڣ���״̬���ڿ�棬�������������壬���ţ�{0}", bundleNo));
                        continue;
                    }

                    // �����ݿ����ҵ�ʯ�����࣬����Ҳ��������ļ���ָ����ʯ�����࣬�򲻵������
                    var category = getStoneCategory(importBundle.Category.Name);
                    if (category == null)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("����Ĵ�����ʯ��������ϵͳ�в����ڣ����ܵ���Ĵ��������ţ�{0}", bundleNo));
                        continue;
                    }
                    else
                    {
                        importBundle.CategoryId = category.Id;
                    }

                    // �����ݿ����ҵ��ȼ���Ϣ������뵥�ĵȼ���ϵͳ�в����ڣ���ʹ�á�δ֪���ȼ�
                    var grade = getStoneGrade(importBundle.Grade.Name);
                    if (grade == null)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д����ĵȼ���ϵͳ�в����ڣ�ʹ�á�δ֪���ȼ����´��������ţ�{0}", bundleNo));
                        grade = getStoneGrade("δ֪");
                    }
                    importBundle.GradeId = grade.Id;

                    // �����ݿ����ҵ�D99-99������򣨽���6/30/17��D99-99��ΪĬ�ϵĵ�������Ŀ�������������Ҫ�Ժ���ģ�������ӿ��������������ļ��ж�ȡ�͵��룩
                    var psa = getProductStockingArea("D", ProductStockingAreaType.AShelf, 99, 99);
                    importBundle.StockingAreaId = psa.Id;

                    importBundle.StockInOperator = dataOperatorName;
                    importBundle.StockInTime = DateTime.Now;
                    await InitializeDBRecord(importBundle);
                    DbContext.StoneBundles.Add(importBundle);

                    // ���¶�Ӧ���ϵ�״̬��������ڻ��ϣ�����ӿ��״̬���µ������������
                    if (block != null && block.Status == BlockStatus.InStock)
                        block.Status = BlockStatus.Processed;

                    addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ�����������ݣ����ţ�{0}", bundleNo));
                }
                else
                {
                    // ���ݿ�����������壬�жϴ���״̬�����Ƿ����
                    bool shouldUpdate = bundleInDB.NotVerified && bundleInDB.Slabs.Count == 0;
                    if (!shouldUpdate)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("����Ĵ������������ݿ����Ѿ��������Ǿ���ϵͳ�������̲����Ĵ��������ܵ������ݣ����ţ�{0}", bundleNo));
                        continue;
                    }

                    if (bundleInDB.Status != SlabBundleStatus.InStock)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("���ݿ��ж�Ӧ�Ĵ��������ڿ��״̬�����ܸ������ݣ����ţ�{0}", bundleNo));
                        continue;
                    }

                    // ���������Ϣ��һ�£��������
                    if (bundleInDB.TotalBundleCount == importBundle.TotalBundleCount &&
                        bundleInDB.TotalSlabCount == importBundle.TotalSlabCount &&
                        bundleInDB.LengthAfterStockIn == importBundle.LengthAfterStockIn &&
                        bundleInDB.WidthAfterStockIn == importBundle.WidthAfterStockIn &&
                        bundleInDB.Category.Name == importBundle.Category.Name &&
                        bundleInDB.Grade.Name == importBundle.Grade.Name &&
                        bundleInDB.Area == importBundle.Area &&
                        bundleInDB.Thickness == importBundle.Thickness)
                    {
                        addInfoMsg(dataUpdatingMsgs, string.Format("����Ĵ������ݺ����ݿ�������һ�£�������£����ţ�{0}", bundleNo));
                        continue;
                    }

                    // �������ʯ�����಻һ�£����´�����ʯ�����࣬����Ҳ��������ļ���ָ����ʯ�����࣬�򲻵������
                    if (bundleInDB.Category.Name != importBundle.Category.Name)
                    {
                        var category = getStoneCategory(importBundle.Category.Name);
                        if (category == null)
                        {
                            addErrorMsg(dataUpdatingMsgs, string.Format("����Ĵ�����ʯ��������ϵͳ�в����ڣ����ܵ���Ĵ��������ţ�{0}", bundleNo));
                            continue;
                        }
                        else
                        {
                            bundleInDB.CategoryId = category.Id;
                        }
                    }

                    // ��������ȼ���һ�£����´����ĵȼ���Ϣ������뵥�ĵȼ���ϵͳ�в����ڣ���ʹ�á�δ֪���ȼ�
                    if (bundleInDB.Grade.Name != importBundle.Grade.Name)
                    {
                        var grade = getStoneGrade(importBundle.Grade.Name);
                        if (grade == null)
                        {
                            addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д����ĵȼ���ϵͳ�в����ڣ�ʹ�á�δ֪���ȼ����´��������ţ�{0}", bundleNo));
                            grade = getStoneGrade("δ֪");
                        }
                        bundleInDB.GradeId = grade.Id;
                    }

                    // ������������
                    bundleInDB.TotalBundleCount = importBundle.TotalBundleCount;
                    bundleInDB.TotalSlabCount = importBundle.TotalSlabCount;
                    bundleInDB.LengthAfterStockIn = importBundle.LengthAfterStockIn;
                    bundleInDB.WidthAfterStockIn = importBundle.WidthAfterStockIn;
                    bundleInDB.Area = importBundle.Area;
                    bundleInDB.Thickness = importBundle.Thickness;
                    bundleInDB.LastUpdatedTime = DateTime.Now;
                    addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ��������ݿ��еĴ������ݣ����ţ�{0}", bundleNo));
                }

                DbContext.SaveChanges();
            }

            return Success(new { DataUpdatingMessages = dataUpdatingMsgs, DataParsingMessages = dataParsingMsgs });
        }

        ProductStockingArea getProductStockingArea(string section, ProductStockingAreaType type, int segment, int slot)
        {
            var psa = DbContext.ProductStockingAreas
                .Where(sa => sa.Section == section && sa.Type == type && sa.Segment == sa.Segment && sa.Slot == slot)
                .SingleOrDefault();
            return psa;
        }

        /// <summary>
        /// �����ݿ���ͨ��ʯ���������ƻ�ȡʯ��������Ϣ
        /// </summary>
        /// <param name="categoryName">ʯ����������</param>
        /// <returns>StoneCategory��������Ҳ����򷵻�null</returns>
        StoneCategory getStoneCategory(string categoryName)
        {
            var category = DbContext.StoneCategories
                            .Where(sc => sc.Name == categoryName)
                            .SingleOrDefault();
            return category;
        }

        /// <summary>
        /// �����ݿ���ͨ��ʯ�ĵȼ����ƻ�ȡʯ�ĵȼ���Ϣ
        /// </summary>
        /// <param name="categoryName">ʯ�ĵȼ�����</param>
        /// <returns>StoneGrade��������Ҳ����򷵻�null</returns>
        StoneGrade getStoneGrade(string gradeName)
        {
            var grade = DbContext.StoneGrades
                            .Where(g => g.Name == gradeName)
                            .SingleOrDefault();
            return grade;
        }

        /// <summary>
        /// �������ĥ���ʼ�����
        /// </summary>
        /// <param name="specFile">����ĥ���ʼ������ļ���Ϣ</param>
        /// <returns>�����������а�������ʹ�����Ϣ</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("ImportPolishingData")]
        [Authorize(Roles = RoleDefinition.DataOperator)]
        public async Task<IActionResult> ImportPolishingData([FromBody] BundleSpecFileViewModel specFile)
        {
            string base64Content = specFile.FileContent;

            if (string.IsNullOrEmpty(specFile.FileName))
                return BadRequest("�ļ�������Ϊ��");

            if (!specFile.FileName.EndsWith(".xlsx"))
                return Error("�ϴ����ļ�������Excel 2007�Ժ��ʽ��.xlsx�ļ�");

            if (string.IsNullOrEmpty(base64Content))
                return BadRequest("�ļ����ݲ���Ϊ��");

            byte[] fileBytes = null;
            try
            {
                fileBytes = Convert.FromBase64String(base64Content);
            }
            catch
            {
                return Error("�ļ����ݲ��Ϸ�������ʧ��");
            }

            List<string> dataParsingMsgs = new List<string>();
            List<string> dataUpdatingMsgs = new List<string>();
            List<StoneBundle> importBundles = new List<StoneBundle>();

            logger.LogTrace("���ļ��ж�ȡ������Ϣ");

            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                PolishedBundleSpecSheetProcessingResult result = bundleSpecSheetSVC.GetPolishedBundleSpecs(ms);
                dataParsingMsgs = result.ErrorMessages;
                importBundles = result.Bundles;
            }

            if (dataParsingMsgs.Count > 0)
                return Success(new { DataUpdatingMessages = dataUpdatingMsgs, DataParsingMessages = dataParsingMsgs });

            string dataOperatorName = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            foreach (StoneBundle importBundle in importBundles)
            {
                string bundleNo = string.Format("{0} #{1}-{2}", importBundle.BlockNumber, importBundle.TotalBundleCount, importBundle.BundleNo);
                logger.LogTrace("�������ĥ������ - {bundleNo}", bundleNo);

                if (importBundle.Slabs.Count == 0)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("�뵥����û���κκϷ��Ĵ�����ݣ����ţ�{0}", bundleNo));
                    continue;
                }

                // ������Ӧ�Ļ�����Ϣ
                var block = DbContext.Blocks
                    .Include(b => b.Bundles)
                        .ThenInclude(sb => sb.Slabs)
                    .Where(b => b.BlockNumber == importBundle.BlockNumber)
                    .SingleOrDefault();

                if (block == null)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("������Ӧ�Ļ�����ϵͳ�в����ڣ����ţ�{0}", bundleNo));
                    continue;
                }

                if (block.Status != BlockStatus.Manufacturing)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("������Ӧ�Ļ��ϲ���������״̬�����ţ�{0}", bundleNo));
                    continue;
                }

                // �ҳ����ݿ��ж�Ӧ�Ĵ���
                StoneBundle bundleInDB = block.Bundles.ToList().Find(sb => sb.BundleNo == importBundle.BundleNo);

                if (bundleInDB == null)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("�뵥�еĴ�����ϵͳ�����ڣ����ţ�{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Slabs.Count == 0)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("�뵥�еĴ�����ϵͳ��û���κδ�����ݣ����ţ�{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Status != SlabBundleStatus.Manufacturing || (bundleInDB.ManufacturingState != ManufacturingState.Completed && bundleInDB.ManufacturingState != ManufacturingState.Filled))
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("������״̬���������׶β��ܵ���ĥ�����ݣ����ţ�{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Thickness != importBundle.Thickness)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д����ĺ�Ⱥ�ϵͳ�в�һ�£�����ϵ����Ա�������ݣ����ţ�{0}", bundleNo));
                }

                // ���´����ĵȼ���Ϣ������뵥�ĵȼ���ϵͳ�в����ڣ���ʹ�á�δ֪���ȼ�
                var grade = getStoneGrade(importBundle.Grade.Name);
                if (grade == null)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д����ĵȼ���ϵͳ�в����ڣ�ʹ�á�δ֪���ȼ����´��������ţ�{0}", bundleNo));
                    grade = getStoneGrade("δ֪");
                }
                bundleInDB.Grade = grade;
                bundleInDB.GradeId = grade.Id;

                // �ҳ��������Ļ��϶�Ӧ�Ĺ���
                // 2017-6-27���޸�BUG #632�����ӹ�����������ɸѡ������ֻ���ҵȴ�ĥ���ʼ�͵ȴ�ĥ��ȷ�ϵĹ���������ͬһ�Ż��ϱ���������������ϵͳ�������BUG #632��
                var wo = DbContext.WorkOrders
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .Where(w => w.MaterialRequisition.Block.BlockNumber == importBundle.BlockNumber && (w.ManufacturingState == ManufacturingState.Filled || w.ManufacturingState == ManufacturingState.PolishingQEFinished))
                    .SingleOrDefault();

                foreach (Slab importSlab in importBundle.Slabs)
                {
                    Slab slab = bundleInDB.Slabs.ToList().Find(s => s.SequenceNumber == importSlab.SequenceNumber);
                    // �뵥�еĴ�������ݿ��Ӧ�Ĵ����в����ڣ������ǣ�1������������������������2���µĴ�壩
                    if (slab == null)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�еĴ����ϵͳ�Ķ�Ӧ�����в����ڣ����Խ�������뵱ǰ�������´�����ݣ����ţ�{0}�������ţ�{1}", bundleNo, importSlab.SequenceNumber));
                        slab = DbContext.Slabs
                            .Include(s => s.Bundle)
                                .ThenInclude(b => b.Slabs)
                            .Include(s => s.Block)
                            .Where(s => s.SequenceNumber == importSlab.SequenceNumber && s.Block.BlockNumber == importBundle.BlockNumber)
                            .SingleOrDefault();

                        // �����ϵͳ�в����ڣ��Զ����µĴ�嵼�뵽���ݿ�
                        if (slab == null)
                        {
                            slab = new Slab()
                            {
                                Block = block,
                                Bundle = bundleInDB,
                                CategoryId = bundleInDB.CategoryId,
                                CreatedTime = DateTime.Now,
                                Creator = dataOperatorName,
                                LastUpdatedTime = DateTime.Now,
                                LengthAfterSawing = importSlab.LengthAfterPolishing.Value,
                                WidthAfterSawing = importSlab.WidthAfterPolishing.Value,
                                ManufacturingState = bundleInDB.ManufacturingState,
                                SawingNote = "ĥ���ʼ�׶��¼Ӵ��",
                                SequenceNumber = importSlab.SequenceNumber,
                                Status = bundleInDB.Status,
                                Thickness = bundleInDB.Thickness,
                                Type = SlabType.Polished
                            };
                            // �������û��������������������Ӳ����ߴ磬�����ò����ߴ�����
                            if (wo != null && wo.SkipFilling == false)
                            {
                                slab.LengthAfterFilling = slab.LengthAfterSawing;
                                slab.WidthAfterFilling = slab.WidthAfterSawing;
                            }
                            addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ����뵥�еĴ����ӵ�ϵͳ�У����ţ�{0}�������ţ�{1}", bundleNo, importSlab.SequenceNumber));
                        }
                        else
                        {
                            // 2017-6-27��James���޸�BUG #636
                            slab.Bundle.Slabs.Remove(slab); // ��ԭʼ�����н�Ҫ�����Ĵ���Ƴ�
                            slab.Bundle.Area = Helper.GetAvailableSlabAreaForBundle(slab.Bundle);   // ���¼���ԭʼ����������ʹ������
                            slab.Bundle.TotalSlabCount = Helper.GetAvailableSlabCount(slab.Bundle);
                            addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ����ô��ӵ�{2}���ƶ�����ǰ������ǰ���ţ�{0}�������ţ�{1}", bundleNo, importSlab.SequenceNumber, slab.Bundle.BundleNo));
                        }

                        // �������ӵ����ݿ�Ĵ�����
                        slab.Bundle = bundleInDB;
                        bundleInDB.Slabs.Add(slab);
                    }

                    if (slab.Status != SlabBundleStatus.Manufacturing || (slab.ManufacturingState != ManufacturingState.Filled && slab.ManufacturingState != ManufacturingState.Completed))
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("�뵥�д���״̬���������׶β��ܵ���ĥ�����ݣ����ţ�{0}�������ţ�{1}", bundleNo, importSlab.SequenceNumber));
                        continue;
                    }

                    slab.LengthAfterPolishing = importSlab.LengthAfterPolishing;
                    slab.WidthAfterPolishing = importSlab.WidthAfterPolishing;
                    slab.DeductedLength = importSlab.DeductedLength;
                    slab.DeductedWidth = importSlab.DeductedWidth;
                    slab.ManufacturingState = ManufacturingState.Completed;
                    slab.PolishingNote = importSlab.PolishingNote;
                    slab.LastUpdatedTime = DateTime.Now;
                    slab.GradeId = bundleInDB.GradeId;

                    addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ����´�����ݣ����ţ�{0}�������ţ�{1}", bundleNo, importSlab.SequenceNumber));
                }

                bundleInDB.Area = Helper.GetAvailableSlabAreaForBundle(bundleInDB);
                bundleInDB.TotalSlabCount = Helper.GetAvailableSlabCount(bundleInDB);
                bundleInDB.LastUpdatedTime = DateTime.Now;

                // �ж��뵥�еĴ�����ϵͳ�еĴ����Ĵ�������Ƿ�һ�£����һ������Կ����Զ���ɴ���ĥ���ʼ�
                // һ�µı�׼�����Ƭ��һ���Լ������еĴ�����һһ��Ӧ
                bool bundleQEAutoCompletionEligible = false;    // �Ƿ�߱��Զ���ɴ���ĥ���ʼ������
                if (bundleInDB.TotalSlabCount != importBundle.TotalSlabCount)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д�������Ƭ����ϵͳ�ж�Ӧ������Ƭ����һ�£������Զ���ɴ�����ĥ���ʼ죬�븴�����ݺ��ֹ���ɴ���ĥ���ʼ죬���ţ�{0}", bundleNo));
                }
                else
                {
                    List<int> slabSeqNumInDB = bundleInDB.Slabs.ToList().FindAll(s => s.Status == SlabBundleStatus.Manufacturing).Select(s => s.SequenceNumber).ToList();
                    List<int> slabSeqNumInFile = importBundle.Slabs.ToList().Select(s => s.SequenceNumber).ToList();

                    // ���Ա����ݿ�����е���Ƭ�����뵥���˹���д�Ĵ�����Ƭ���ǲ����ģ���Ҫ���߷��������Ĵ����Ƭ�����ݣ���ʵ���Ƭ��)
                    // ����������������������ݿ��е���Ƭ�����뵥���˹���д��Ƭ��һ�£��������뵥�д�������������⵼�´��û�б���ӵ��ڴ��еĴ�����
                    // �����������ʱ��ѭ���Ա����������Ĵ����ſ��ܻᱨ��DB�д�����Ƭ�������ڴ��е���Ĵ�����ʵ�ʴ��Ƭ������OutOfRangeException��
                    if (slabSeqNumInDB.Count != slabSeqNumInFile.Count)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д�������Ƭ����ϵͳ�ж�Ӧ������Ƭ����һ�£������Զ���ɴ�����ĥ���ʼ죬�븴�����ݺ��ֹ���ɴ���ĥ���ʼ죬���ţ�{0}", bundleNo));
                    }
                    else
                    {
                        slabSeqNumInDB.Sort();
                        slabSeqNumInFile.Sort();
                        bool slabSeqNumMatched = true;
                        for (int i = 0; i < slabSeqNumInDB.Count; i++)
                        {
                            if (slabSeqNumInDB[i] != slabSeqNumInFile[i])
                            {
                                slabSeqNumMatched = false;
                                addWarningMsg(dataUpdatingMsgs, string.Format("�뵥�д����Ĵ����ź�ϵͳ�д����Ĵ�����û��һһ��Ӧ�������Զ���ɴ�����ĥ���ʼ죬�븴�����ݺ��ֹ���ɴ���ĥ���ʼ죬���ţ�{0}", bundleNo));
                                break;
                            }
                        }
                        bundleQEAutoCompletionEligible = slabSeqNumMatched;
                    }
                }

                // ����߱��Զ���ɴ���ĥ���ʼ����������������µ����д��״̬��������״̬Ϊ���ϻ���Completed����Ѵ���״̬Ҳ��ΪCompleted
                if (bundleQEAutoCompletionEligible)
                {
                    bool allSlabQEFinished = bundleInDB.Slabs.All(s => s.ManufacturingState == ManufacturingState.Completed || s.Status == SlabBundleStatus.Discarded);
                    if (allSlabQEFinished)
                    {
                        bundleInDB.ManufacturingState = ManufacturingState.Completed;
                        addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ���ɴ���ĥ���ʼ죬���ţ�{0}", bundleNo));
                    }
                }

                // ���¹�����ͳ������
                if (wo != null)
                {
                    wo.AreaAfterPolishing = Helper.GetBlockManufacturingArea(block, ManufacturingState.Completed);
                    wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
                    wo.LastUpdatedTime = DateTime.Now;
                }

                //���ÿŻ��������д���״̬�����ȫ����ΪCompleted����ѹ���״̬��Ϊ��һ������״̬�����¹������ʲ��ҷ���֪ͨ
                bool allBundleQEFinished = block.Bundles.All(sb => sb.ManufacturingState == ManufacturingState.Completed);
                if (allBundleQEFinished)
                {
                    wo.ManufacturingState = ManufacturingState.PolishingQEFinished;

                    wo.PolishedSlabOutturnPercentage = wo.AreaAfterPolishing.Value / Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value);
                    wo.PolishedSlabOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.PolishedSlabOutturnPercentage.Value);

                    addInfoMsg(dataUpdatingMsgs, string.Format("�ɹ���ɹ�����ĥ���ʼ칤�򣬻��Ϻţ�{0}�������ţ�{1}", block.BlockNumber, wo.OrderNumber));

                    // ��ĥ�����ܷ��Ͷ�����Ϣ
                    string polishingInfoPageUrl = "/workOrders/info/" + wo.Id;
                    string title = string.Format("ĥ���ʼ��������ύ�������ĥ��ȷ��\n���ϱ�ţ�{0} ", block.BlockNumber);
                    string text = string.Format("���� {0} ��ĥ���ʼ��������ύ�������ĥ��ȷ��", wo.OrderNumber);
                    await SendDingtalkLinkMessage(title, text, polishingInfoPageUrl, RoleDefinition.SlabPolishingManager, specFile.AuthCode);
                }

                DbContext.SaveChanges();
            }

            return Success(new { DataUpdatingMessages = dataUpdatingMsgs, DataParsingMessages = dataParsingMsgs });
        }

        void addErrorMsg(List<string> msgs, string msg)
        {
            logger.LogError(msg);
            addMsg(msgs, string.Format("����{0}", msg));
        }
        void addWarningMsg(List<string> msgs, string msg)
        {
            logger.LogWarning(msg);
            addMsg(msgs, string.Format("���棺{0}", msg));
        }
        void addInfoMsg(List<string> msgs, string msg)
        {
            logger.LogInformation(msg);
            addMsg(msgs, msg);
        }

        void addMsg(List<string> msgs, string msg)
        {
            msgs.Add(msg);
        }

        private List<ManufacturingState> getRoleAssignedManufacturingState(IList<string> roles)
        {
            List<ManufacturingState> states = new List<ManufacturingState>();
            roles.ToList().ForEach(r =>
            {
                var roleStates = getRoleAssigedManufacturingState(r);
                roleStates.ForEach(rs =>
                {
                    if (!states.Exists(s => { return s == rs; }))
                        states.Add(rs);
                });
            });
            return states;
        }

        private List<ManufacturingState> getRoleAssigedManufacturingState(string role)
        {

            var roleAssignments = Helper.RoleAssignedStates;
            if (roleAssignments.Keys.Contains(role))
                return roleAssignments[role].ToList();
            else
                return new List<ManufacturingState>();
        }
    }
}
