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
    /// 生产工单API
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
        /// 获取所有生产工单，提供分页机制
        /// </summary>
        /// <param name="pageSize">每页多少条数据，如不提供则默认每页20条</param>
        /// <param name="pageNo">页号，如不提供则默认第1页</param>
        /// <param name="blockNumber">可选，荒料编号，如果不提供，则表示所有荒料</param>
        /// <param name="orderNumber">可选，工单编号，如果不提供，则表示所有工单</param>
        /// <param name="orderNumberforSO">可选，订单编号，如果不提供，则表示所有订单</param>
        /// <param name="statusCodes">可选，销售订单状态，多个订单状态使用英文逗号分隔。如果不提供，则表示所有状态</param>
        /// <returns></returns>
        [HttpGet("pageSize,pageNo,stateCodes,blockNumber,orderNumber,orderNumberforSO")]
        [Route("GetAll")]
        public async Task<IActionResult> GetAll([FromQuery] List<ushort> stateCodes, [FromQuery] string blockNumber, [FromQuery] string orderNumber, [FromQuery] string orderNumberforSO, [FromQuery] int? pageSize = 20, [FromQuery] int? pageNo = 1)
        {
            if (pageSize <= 0 || pageNo <= 0)
                return BadRequest();

            if (!isStatusCodesValid(stateCodes))
                return BadRequest("存在不合法的状态码");

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
        /// 获取一个生产工单的信息
        /// </summary>
        /// <param name="id">Id</param>
        /// <returns>如果成功，则返回标准Success数据结构，content属性里包含生产工单信息，如果找不到对应Id的生产工单，返回标准Error数据结构，其他失败返回相应的HTTP错误代码</returns>
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
                return Error("找不到指定的生产工单");

            await Helper.HiddenAmount(order.SalesOrder, UserManager, User);

            return Success(order);
        }

        /// <summary>
        /// 新建生产工单
        /// </summary>
        /// <param name="workOrder">荒料修边JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("工单号已存在");

            if (workOrder.Priority == null)
                return BadRequest("请输入工单优先级");

            var so = DbContext.SalesOrders
                    .Where(s => s.Id == workOrder.SalesOrderId)
                    .Include(d => d.Details)
                    .SingleOrDefault();

            if (so == null)
                return BadRequest("销售订单不存在");

            var sod = DbContext.SalesOrderDetails
                    .Where(d => d.Id == workOrder.SalesOrderDetailId)
                    .SingleOrDefault();

            if (sod == null)
                return BadRequest("销售明细不存在");

            if (sod.OrderId != so.Id)
                return BadRequest("销售订单与销售明细不匹配");

            if (so.OrderType == SalesOrderType.BlockNotInStock || so.OrderType == SalesOrderType.BlockInStock || so.OrderType == SalesOrderType.PolishedSlabInStock)
                return BadRequest("销售有库存荒料或无库存荒料或者有库存光板都不需要新建工单");

            if (so.Status != SalesOrderStatus.Approved)
                return BadRequest("该销售订单的状态不允许新建生产工单");

            WorkOrder dbWorkOrder = mapper.Map<WorkOrder>(workOrder);
            dbWorkOrder.Thickness = sod.Specs.Object.Height;
            dbWorkOrder.DeliveryDate = workOrder.DeliveryDate.Value.ToLocalTime();
            await base.InitializeDBRecord(dbWorkOrder);

            DbContext.WorkOrders.Add(dbWorkOrder);
            DbContext.SaveChanges();

            // 给各主管和仓管发消息
            string workOrderInfoPageUrl = "/workOrders/info/" + dbWorkOrder.Id;
            string title = "新的工单已创建，请安排生产";
            string text = string.Format("工单 {0} 已提交，请安排生产", dbWorkOrder.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderInfoPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, workOrder.AuthCode);

            return Success();
        }

        /// <summary>
        /// 更新生产工单
        /// </summary>
        /// <param name="workOrder">更新生产工单数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("工单不存在");

            if (workOrder.Priority == null)
                return BadRequest("请输入工单优先级");

            if (workOrder.DeliveryDate.Value.ToLocalTime() < DateTime.Today)
                return BadRequest("交付日期不能是过去的日期");

            if (wo.ManufacturingState == ManufacturingState.Completed || wo.ManufacturingState == ManufacturingState.Cancelled)
                return BadRequest("生产工单状态不允许修改工单信息");

            if (wo.OrderType != workOrder.OrderType)
            {
                if (wo.ManufacturingState == ManufacturingState.NotStarted || wo.ManufacturingState == ManufacturingState.MaterialRequisitionSubmitted
                || wo.ManufacturingState == ManufacturingState.MaterialRequisitioned || wo.ManufacturingState == ManufacturingState.TrimmingDataSubmitted
                || wo.ManufacturingState == ManufacturingState.Trimmed || wo.ManufacturingState == ManufacturingState.SawingDataSubmitted
                || wo.ManufacturingState == ManufacturingState.Sawed)
                    wo.OrderType = workOrder.OrderType;
                else
                    return BadRequest("此生产工单的状态不允许修改生产类型");
            }

            wo.Priority = workOrder.Priority;
            wo.DeliveryDate = workOrder.DeliveryDate.Value.ToLocalTime();
            wo.LastUpdatedTime = DateTime.Now;
            wo.Notes = workOrder.Notes;

            DbContext.SaveChanges();
            // 给各主管和仓管发消息
            string workOrderInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = "工单内容有更新";
            string text = string.Format("工单 {0} 有更新，请查看并对生产计划做相应调整", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderInfoPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, workOrder.AuthCode);
            return Success();
        }

        /// <summary>
        /// 获取生产工单的领料信息
        /// </summary>
        /// <param name="workOrderId">工单Id</param>
        /// <returns>如果成功，则返回标准Success数据结构，content属性里包含领料信息，如果找不到对应Id的生产工单，返回标准Error数据结构，其他失败返回相应的HTTP错误代码</returns>
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
        /// 添加领料单
        /// </summary>
        /// <param name="mateReq">领料单Json数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("AddMaterialRequisition")]
        public async Task<IActionResult> AddMaterialRequisition([FromBody] MaterialRequisitionViewModel mateReq)
        {
            if (mateReq.WorkOrderId <= 0)
                return BadRequest("工单Id不合法");

            var wo = DbContext.WorkOrders
                .Where(w => w.Id == mateReq.WorkOrderId)
                .Include(w => w.SalesOrderDetail)
                .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单不存在");

            if (wo.ManufacturingState != ManufacturingState.NotStarted)
                return BadRequest("只有尚未开始生产的工单才允许添加领料单");

            if (mateReq.BlockId == null && mateReq.BundleId == null)
                return BadRequest("领料单必须指定荒料Id或者大板扎号Id");

            Block block = null;
            if (mateReq.BlockId != null)
            {
                int blockId = mateReq.BlockId.Value;
                block = DbContext.Blocks
                    .Where(b => b.Id == blockId)
                    .SingleOrDefault();

                if (block == null)
                    return Error("领料单中指定的荒料不存在，荒料Id" + blockId);

                var stoneCategory = DbContext.StoneCategories
                    .Include(c => c.BaseCategory)
                    .Where(sc => sc.Id == wo.SalesOrderDetail.Specs.Object.CategoryId)
                    .AsNoTracking()
                    .SingleOrDefault();

                // 确保领取荒料的石材种类是销售明细中指定的种类或者其种类的基础荒料种类，如不是则报错
                if (block.CategoryId != stoneCategory.Id && (stoneCategory.BaseCategory == null || stoneCategory.BaseCategory.Id != block.CategoryId))
                    return Error("领料单中指定的荒料与订单需求不一致，荒料Id" + blockId);

                if (block.Status != BlockStatus.InStock)
                    return Error("领料单中指定的荒料不是在库状态，荒料Id" + blockId);

                block.Status = BlockStatus.Reserved;
            }

            //todo: 添加代码进行大板扎的状态检查并且进行状态更新。目前还存在问题：领料单可以领取多扎大板，所以需要改现在领料单中存储大板扎的数据结构

            var dbMR = mapper.Map<MaterialRequisition>(mateReq);
            await InitializeDBRecord(dbMR);

            wo.MaterialRequisition = dbMR;
            wo.ManufacturingState = ManufacturingState.MaterialRequisitionSubmitted;
            DbContext.SaveChanges();

            // 给审批人发送钉钉消息
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
            {
                string approvalPageUrl = "/workOrders/info/" + dbMR.Id;
                string title = string.Format("新的领料单需要审批\n荒料编号：{0} ", wo.MaterialRequisition.Block.BlockNumber);
                string text = string.Format("工单 {0} 的领料单已提交，请审批", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, approvalPageUrl, RoleDefinition.BlockManager, mateReq.AuthCode);
            }
            else
            {
                string approvalPageUrl = "/workOrders/info/" + dbMR.Id;
                string title = string.Format("新的领料单需要审批\n荒料编号：{0} ", wo.MaterialRequisition.Block.BlockNumber);
                string text = string.Format("工单 {0} 的领料单已提交，请审批", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, approvalPageUrl, RoleDefinition.ProductManager, mateReq.AuthCode);
            }

            return Success();
        }

        /// <summary>
        /// 审批领料申请
        /// </summary>
        /// <param name="workOrderId">生产工单Id</param>
        /// <param name="authCode">发送钉钉消息验证码</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("工单不存在");

            if (wo.ManufacturingState == ManufacturingState.MaterialRequisitioned)
                return Success();

            if (wo.ManufacturingState != ManufacturingState.MaterialRequisitionSubmitted)
                return BadRequest("该工单的生产状态不允许批准领料申请");

            if (wo.MaterialRequisition == null)
                return BadRequest("该工单下没有提交领料单");

            if ((wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab) && wo.MaterialRequisition.Block == null)
                return BadRequest("该工单下的领料单没有荒料信息");

            if (wo.OrderType == WorkOrderType.Tile && wo.MaterialRequisition.Bundle == null)
                return BadRequest("该工单下的领料单没有大板信息");

            // 更新工单状态
            wo.ManufacturingState = ManufacturingState.MaterialRequisitioned;

            // 更新对应的材料的状态
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
                wo.MaterialRequisition.Block.Status = BlockStatus.Manufacturing;

            if (wo.OrderType == WorkOrderType.Tile)
                wo.MaterialRequisition.Bundle.Status = SlabBundleStatus.CheckedOut;

            Block block = wo.MaterialRequisition.Block;

            block.StockOutOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            block.StockOutTime = DateTime.Now;

            DbContext.SaveChanges();

            // 给领料人发送钉钉消息
            if (wo.OrderType == WorkOrderType.PolishedSlab || wo.OrderType == WorkOrderType.RawSlab)
            {
                // 给大锯主管发送钉钉消息
                string trimmingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("领料单已审批，请联系库管领料\n荒料编号：{0} ", block.BlockNumber);
                string text = string.Format("工单 {0} 的领料单已审批，请联系库管领料", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, trimmingInfoPageUrl, RoleDefinition.SawingManager, authCode);
            }
            else
            {
                // 给红外线主管发送钉钉消息 //todo: 现在没有红外线组主管的角色
                string trimmingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("领料单已审批，请联系库管领料\n荒料编号：{0} ", block.BlockNumber);
                string text = string.Format("工单 {0} 的领料单已审批，请联系库管领料", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, trimmingInfoPageUrl, RoleDefinition.TileQE, authCode);
            }

            return Success();
        }

        /// <summary>
        /// 荒料修边
        /// </summary>
        /// <param name="trimmingInfo">荒料修边JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("UpdateTrimmingData")]
        public async Task<IActionResult> UpdateTrimmingData([FromBody] TrimmingInfo trimmingInfo)
        {
            if (trimmingInfo.WorkOrderId <= 0)
                return BadRequest("工单号Id不合法");

            if (trimmingInfo.TrimmingStartTime > trimmingInfo.TrimmingEndTime)
                return BadRequest("修边开始时间不能晚于修边结束时间");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == trimmingInfo.WorkOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单不存在");

            if (wo.ManufacturingState != ManufacturingState.MaterialRequisitioned)
                return BadRequest("该工单的状态不允许修边");

            if (trimmingInfo.TrimmedHeight <= 0 || trimmingInfo.TrimmedLength <= 0 || trimmingInfo.TrimmedWidth <= 0)
                return BadRequest("修边尺寸不合法");

            if (wo.OrderType != WorkOrderType.RawSlab && wo.OrderType != WorkOrderType.PolishedSlab)
                return BadRequest("工单类型不支持修边操作");

            if (wo.MaterialRequisition == null)
                return BadRequest("该工单没有对应的领料单，不能进行修边操作");

            if (wo.MaterialRequisition.BlockId == null)
                return BadRequest("该工单的领料单中没有对应的荒料信息");

            Block block = wo.MaterialRequisition.Block;

            if (block == null)
                return BadRequest("该工单领料单中指定的荒料不存在");

            if (block.Status != BlockStatus.Manufacturing)
                return BadRequest("荒料状态不允许修边");

            wo.ManufacturingState = ManufacturingState.TrimmingDataSubmitted;
            wo.TrimmingDetails = trimmingInfo.TrimmingDetails;
            wo.TrimmingStartTime = trimmingInfo.TrimmingStartTime.Value.ToLocalTime();
            wo.TrimmingEndTime = trimmingInfo.TrimmingEndTime.Value.ToLocalTime();
            block.TrimmedHeight = trimmingInfo.TrimmedHeight;
            block.TrimmedLength = trimmingInfo.TrimmedLength;
            block.TrimmedWidth = trimmingInfo.TrimmedWidth;
            wo.TrimmingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // 给修边质检发送钉钉消息
            string trimmingQEPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("修边数据已提交，请进行修边质检\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的修边数据已提交，请对修边后的荒料进行质检", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, trimmingQEPageUrl, RoleDefinition.TrimmingQE, trimmingInfo.AuthCode);

            return Success();
        }

        /// <summary>
        /// 荒料修边质检员确认修边数据API
        /// </summary>
        /// <param name="trimmingQE">荒料JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("指定Id的生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.TrimmingDataSubmitted)
                return BadRequest("此荒料对应工单的状态不允许修边");

            if (trimmingQE.TrimmedHeight <= 0 || trimmingQE.TrimmedLength <= 0 || trimmingQE.TrimmedWidth <= 0)
                return BadRequest("修边尺寸不合法");

            if (wo.OrderType != WorkOrderType.RawSlab && wo.OrderType != WorkOrderType.PolishedSlab)
                return BadRequest("工单类型不支持修边操作");

            if (wo.MaterialRequisition == null)
                return BadRequest("该工单没有对应的领料单，不能进行修边操作");

            if (wo.MaterialRequisition.BlockId == null)
                return BadRequest("该工单的领料单中没有对应的荒料信息");

            int blockId = wo.MaterialRequisition.BlockId.Value;

            var block = DbContext.Blocks
                .Where(b => b.Id == blockId)
                .SingleOrDefault();

            if (block == null)
                return BadRequest("该工单领料单中指定的荒料不存在");

            if (block.Status != BlockStatus.Manufacturing)
                return BadRequest("荒料状态不允许修边");

            block.TrimmedHeight = trimmingQE.TrimmedHeight;
            block.TrimmedLength = trimmingQE.TrimmedLength;
            block.TrimmedWidth = trimmingQE.TrimmedWidth;
            wo.BlockOutturnPercentage = Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value) / Helper.CalculateVolume(block.QuarryReportedLength, block.QuarryReportedWidth, block.QuarryReportedHeight);
            wo.BlockOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.BlockOutturnPercentage.Value);
            wo.TrimmingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
            wo.ManufacturingState = ManufacturingState.Trimmed;

            DbContext.SaveChanges();

            // 给大锯主管发送钉钉消息
            string sawingInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("修边工序已完成，请进行大切工序\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的修边工序已完成，请进行大切工序", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, sawingInfoPageUrl, RoleDefinition.SawingManager, trimmingQE.AuthCode);
            return Success();
        }

        /// <summary>
        /// 荒料大锯
        /// </summary>
        /// <param name="sawingInfo">荒料大锯JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Sawing")]
        public async Task<IActionResult> Sawing([FromBody] SawingInfo sawingInfo)
        {
            if (sawingInfo.WorkOrderId <= 0)
                return BadRequest("工单号Id不合法");

            if (sawingInfo.SawingStartTime > sawingInfo.SawingEndTime)
                return BadRequest("大锯开始时间不能晚于大锯结束时间");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == sawingInfo.WorkOrderId)
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单不存在");

            if (wo.ManufacturingState != ManufacturingState.Trimmed)
                return BadRequest("该工单的状态不允许大锯");

            wo.ManufacturingState = ManufacturingState.SawingDataSubmitted;
            wo.SawingDetails = sawingInfo.SawingDetails;
            wo.SawingStartTime = sawingInfo.SawingStartTime.Value.ToLocalTime();
            wo.SawingEndTime = sawingInfo.SawingEndTime.Value.ToLocalTime();
            wo.SawingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // 给大锯质检员发送钉钉消息
            string sawingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("大锯数据已提交，请进行毛板质检分扎\n荒料编号：{0} ", wo.MaterialRequisition.Block.BlockNumber);
            string text = string.Format("工单 {0} 的大锯数据已提交，请进行毛板质检分扎", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, sawingQAPageUrl, RoleDefinition.SawingQE, sawingInfo.AuthCode);

            return Success();
        }

        /// <summary> 
        /// 分扎编号
        /// </summary>
        /// <param name="inputSB">大板JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("SplitBundle")]
        public async Task<IActionResult> SplitBundle([FromBody] StoneBundleSplitingInfo inputSB)
        {
            if (inputSB.TotalSlabCount <= 0)
                return BadRequest("总片数不合法");

            if (inputSB.TotalBundleCount <= 0)
                return BadRequest("总扎数不合法");

            if (inputSB.Thickness <= 0)
                return BadRequest("大板厚度不合法");

            var wo = DbContext.WorkOrders
                     .Where(w => w.Id == inputSB.WorkOrderId)
                     .Include(workOrder => workOrder.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                            .ThenInclude(b => b.Bundles)
                     .Include(w => w.SalesOrderDetail)
                     .SingleOrDefault();

            if (wo == null)
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.SawingDataSubmitted)
                return BadRequest("此份生产工单的状态不允许大切");

            if (wo.MaterialRequisition == null)
                return BadRequest("此工单的领料单不存在");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("领料单中的荒料不存在");

            if (wo.SalesOrderDetail == null)
                return BadRequest("生产工单对应的销售明细不存在");

            if (wo.SalesOrderDetail.Specs.Object == null)
                return BadRequest("生产工单对应的销售明细详细信息不存在");

            if (inputSB.Bundles.Count == 0)
                return BadRequest("请输入分扎明细");

            if (inputSB.Bundles.Count != inputSB.TotalBundleCount)
                return BadRequest("输入的扎数与总扎数不相等");

            int j = 0;
            int i = 0;

            foreach (BundleInfo b in inputSB.Bundles)
            {
                j++;

                if (b.BundleNo <= 0 || b.BundleNo > inputSB.TotalBundleCount)
                    return BadRequest("扎号不合法");

                if (b.BundleNo != j)
                    return BadRequest("请按顺序编辑扎号");

                if (b.GradeId <= 0)
                    return BadRequest("石材种类Id不合法");

                var gra = DbContext.StoneGrades
                          .Where(c => c.Id == b.GradeId)
                          .SingleOrDefault();

                if (gra == null)
                    return BadRequest("石材等级不存在");

                foreach (SlabInfo s in b.Slabs)
                {
                    i++;

                    if (s.SequenceNumber <= 0 || s.SequenceNumber > inputSB.TotalSlabCount)
                        return BadRequest("大板编号不合法");

                    if (s.SequenceNumber != i)
                        return BadRequest("请按顺序进行大板编号");

                    if (s.Length <= 0 || s.Width <= 0)
                        return BadRequest("大切后的尺寸不合法");

                    if (s.DeductedLength < 0 || s.DeductedWidth < 0 || s.DeductedLength >= s.Length || s.DeductedWidth >= s.Width)
                        return BadRequest("扣尺尺寸不合法");

                    if (s.Discarded)
                        s.DiscardedReason = s.DiscardedReason.Trim();

                    if (s.Discarded && string.IsNullOrEmpty(s.DiscardedReason))
                        return BadRequest("请输入报废原因");
                }
            }

            var blo = wo.MaterialRequisition.Block;

            if (blo == null)
                return BadRequest("该工单领料单中指定的荒料不存在");

            if (blo.Bundles.Count > 0)
                return BadRequest("该工单领料单中指定的荒料已生产大板");

            if (blo.Status != BlockStatus.Manufacturing)
                return BadRequest("荒料状态不允许大切");

            blo.TotalSlabNo = inputSB.TotalSlabCount;

            // 成品石材种类，在生成的大扎和大板中使用销售明细中的石材种类，而不是荒料的石材种类
            // 因为销售订单明细中可能会由顺切和反切的石材种类
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
            // 直磨，给磨抛质检发送钉钉消息
            if (inputSB.SkipFilling)
            {
                string PolishingingQAPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("大切工序已完成，此颗荒料的大板无需补胶，请进行磨抛工序\n荒料编号：{0} ", blo.BlockNumber);
                string text = string.Format("工单 {0} 的大切工序已完成，此颗荒料的大板无需补胶，请进行磨抛工序", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, PolishingingQAPageUrl, RoleDefinition.PolishingQE, inputSB.AuthCode);

            }
            // 需进行补胶，给补胶主管发送钉钉消息
            else
            {
                string fillingInfoPageUrl = "/workOrders/info/" + wo.Id;
                string title = string.Format("大切工序已完成，此颗荒料的大板无需补胶，请进行磨抛工序\n荒料编号：{0} ", blo.BlockNumber);
                string text = string.Format("工单 {0} 的大切工序已完成，请进行补胶", wo.OrderNumber);
                await SendDingtalkLinkMessage(title, text, fillingInfoPageUrl, RoleDefinition.FillingManager, inputSB.AuthCode);

            }

            return Success();
        }

        /// <summary>
        /// 大板补胶
        /// </summary>
        /// <param name="fillingInfo">大板补胶JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Filling")]
        public async Task<IActionResult> Filling([FromBody] FillingInfo fillingInfo)
        {
            if (fillingInfo.WorkOrderId <= 0)
                return BadRequest("工单号Id不合法");

            if (fillingInfo.FillingStartTime > fillingInfo.FillingEndTime)
                return BadRequest("补胶开始时间不能晚于补胶结束时间");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == fillingInfo.WorkOrderId)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单不存在");

            if (wo.ManufacturingState != ManufacturingState.Sawed && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("该工单的状态不允许补胶");

            if (wo.MaterialRequisition == null)
                return BadRequest("此生产工单的领料单不存在");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("此领料单的荒料不存在");

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

            // 给补胶质检发送钉钉消息
            string fillingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("补胶数据已提交，请进行补胶质检\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的补胶数据已提交，请进行补胶质检", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, fillingQAPageUrl, RoleDefinition.FillingQE, fillingInfo.AuthCode);

            return Success();
        }

        /// <summary>
        /// 大板磨抛
        /// </summary>
        /// <param name="polishingInfo">荒料修边JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Polishing")]
        public async Task<IActionResult> Polishing([FromBody] PolishingInfo polishingInfo)
        {
            if (polishingInfo.WorkOrderId <= 0)
                return BadRequest("工单号Id不合法");

            if (polishingInfo.PolishingStartTime > polishingInfo.PolishingEndTime)
                return BadRequest("磨抛开始时间不能晚于磨抛结束时间");

            var wo = DbContext.WorkOrders
                    .Where(w => w.Id == polishingInfo.WorkOrderId)
                        .Include(w => w.MaterialRequisition)
                            .ThenInclude(mr => mr.Block)
                                .ThenInclude(b => b.Bundles)
                                    .ThenInclude(sb => sb.Slabs)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单不存在");

            if (wo.ManufacturingState != ManufacturingState.PolishingQEFinished)
                return BadRequest("该工单的状态不允许磨抛");

            if (wo.MaterialRequisition == null)
                return BadRequest("此生产工单的领料单不存在");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("此领料单的荒料不存在");

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
                            return BadRequest("此大板状态不允许完成工单");
                    }
                }
                else
                    return BadRequest("此扎大板状态不允许完成工单");
            }

            block.Status = BlockStatus.Processed;
            wo.ManufacturingState = ManufacturingState.Completed;
            wo.PolishingDetails = polishingInfo.PolishingDetails;
            wo.PolishingStartTime = polishingInfo.PolishingStartTime.Value.ToLocalTime();
            wo.PolishingEndTime = polishingInfo.PolishingEndTime.Value.ToLocalTime();
            wo.PolishingOperator = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();

            // 给成品库管发送钉钉消息
            string productStockingInPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("光板已生产完成，请进行光板入库\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的光板已生产完成，请进行光板入库", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, productStockingInPageUrl, new string[] { RoleDefinition.ProductManager, RoleDefinition.PackagingManger, RoleDefinition.FactoryManager }, polishingInfo.AuthCode);
            return Success();
        }

        /// <summary>
        /// 补胶质检
        /// </summary>
        /// <param name="fillingQE">大板JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.FillingDataSubmitted && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单的状态不允许补胶");

            if (fillingQE.Length <= 0 || fillingQE.Width <= 0)
                return BadRequest("补胶尺寸不合法");

            if (fillingQE.DeductedLength < 0 || fillingQE.DeductedWidth < 0 || fillingQE.DeductedLength >= fillingQE.Length || fillingQE.DeductedWidth >= fillingQE.Width)
                return BadRequest("扣尺尺寸不合法");

            var so = wo.SalesOrder;

            if (so == null)
                return BadRequest("销售订单不存在");

            if (so.OrderType != SalesOrderType.PolishedSlabNotInStock)
                return BadRequest("此销售订单类型无需补胶工序");

            var sod = wo.SalesOrderDetail;

            if (sod == null)
                return BadRequest("销售订单明细不存在");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("此工单对应的领料单不存在");

            var block = mr.Block;

            if (block == null)
                return BadRequest("领料单中的荒料不存在");

            if (fillingQE.Discarded)
                fillingQE.DiscardedReason = fillingQE.DiscardedReason.Trim();

            if (fillingQE.Discarded && string.IsNullOrEmpty(fillingQE.DiscardedReason))
                return BadRequest("请输入报废原因");

            var slab = DbContext.Slabs
                      .Where(s => s.Id == fillingQE.SlabId)
                      .Include(s => s.Bundle)
                        .ThenInclude(b => b.Slabs)
                      .SingleOrDefault();

            if (slab == null)
                return BadRequest("大板不存在");

            if (slab.Bundle == null)
                return BadRequest("扎号不存在");

            var stoneBundles = block.Bundles;

            if (stoneBundles.Count <= 0)
                return BadRequest("领料单中的荒料未生产出大板");

            if (!(stoneBundles.Contains(slab.Bundle)))
                return BadRequest("大板不在此份生产工单生产出来的大板中");

            if (slab.Status != SlabBundleStatus.Discarded && slab.Status != SlabBundleStatus.Manufacturing)
                return BadRequest("大板状态不允许补胶");

            if (slab.ManufacturingState != ManufacturingState.FillingDataSubmitted && slab.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("大板生产状态不允许补胶");

            slab.Status = fillingQE.Discarded ? SlabBundleStatus.Discarded : SlabBundleStatus.Manufacturing;
            slab.DiscardedReason = fillingQE.Discarded ? fillingQE.DiscardedReason : null;
            slab.ManufacturingState = fillingQE.Discarded ? ManufacturingState.FillingDataSubmitted : ManufacturingState.Filled;

            slab.LengthAfterFilling = fillingQE.Length;//补胶过程如需提交单片数据，收方的尺寸需长扣5cm,宽扣3cm
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
        /// 补胶质检确认工序结束
        /// </summary>
        /// <param name="workOrderId">生产工单IdJSON数据</param>
        /// <param name="authCode">发送钉钉消息验证码</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.FillingDataSubmitted && wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单的状态不允许补胶");

            if (wo.MaterialRequisition == null)
                return BadRequest("此生产工单的领料单不存在");

            if (wo.MaterialRequisition.Block == null)
                return BadRequest("此领料单的荒料不存在");

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
                        // 如果补胶后长度和宽度为空，表示该块大板没有被补胶质检单片提交过数据，所以补胶结束时要对所有的状态为补胶数据已提交的大板做补胶尺寸的初始化
                        // 将补胶尺寸初始化为大切后的尺寸，对于有补胶数据的大板则不作任何补胶尺寸的更改
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
            // 给磨抛质检发送钉钉消息
            string polishingQAPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("补胶工序已完成，请进行磨抛工序\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的补胶工序已完成，请进行磨抛工序", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, polishingQAPageUrl, RoleDefinition.PolishingQE, authCode);

            return Success();
        }

        /// <summary>
        /// 磨抛质检
        /// </summary>
        /// <param name="polishingQE">大板JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单的状态不允许磨抛");

            if (polishingQE.Length <= 0 || polishingQE.Width <= 0)
                return BadRequest("磨抛尺寸不合法");

            if (polishingQE.DeductedLength < 0 || polishingQE.DeductedWidth < 0 || polishingQE.DeductedLength >= polishingQE.Length || polishingQE.DeductedWidth >= polishingQE.Width)
                return BadRequest("扣尺尺寸不合法");

            var so = wo.SalesOrder;

            if (so == null)
                return BadRequest("销售订单不存在");

            if (so.OrderType != SalesOrderType.PolishedSlabNotInStock)
                return BadRequest("此销售订单类型无需磨抛工序");

            var sod = wo.SalesOrderDetail;

            if (sod == null)
                return BadRequest("销售订单明细不存在");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("此工单对应的领料单不存在");

            var block = mr.Block;

            if (block == null)
                return BadRequest("领料单中的荒料不存在");

            if (polishingQE.Discarded)
                polishingQE.DiscardedReason = polishingQE.DiscardedReason.Trim();

            if (polishingQE.Discarded && string.IsNullOrEmpty(polishingQE.DiscardedReason))
                return BadRequest("请输入报废原因");

            var slab = DbContext.Slabs
                      .Where(s => s.Id == polishingQE.SlabId)
                      .Include(s => s.Bundle)
                        .ThenInclude(b => b.Slabs)
                      .SingleOrDefault();

            if (slab == null)
                return BadRequest("大板不存在");

            if (slab.Bundle == null)
                return BadRequest("扎号不存在");

            var stoneBundles = block.Bundles;

            if (stoneBundles.Count <= 0)
                return BadRequest("领料单中的荒料未生产出大板");

            if (!(stoneBundles.Contains(slab.Bundle)))
                return BadRequest("大板不在此份生产工单生产出来的大板中");

            if (slab.Status != SlabBundleStatus.Manufacturing && slab.Status != SlabBundleStatus.Discarded)
                return BadRequest("大板状态不允许磨抛");

            if (slab.ManufacturingState != ManufacturingState.Filled && slab.ManufacturingState != ManufacturingState.Completed)
                return BadRequest("大板生产状态不允许磨抛");

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
        /// 磨抛后对每扎大板定等级
        /// </summary>
        /// <param name="bundleGradeQE">大板JSON数据</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单的状态不允许磨抛");

            var bundle = DbContext.StoneBundles
                         .Where(b => b.Id == bundleGradeQE.BundleId)
                         .Include(b => b.Slabs)
                         .SingleOrDefault();

            if (bundle == null)
                return BadRequest("此扎大板不存在");

            if (bundle.Slabs.Count <= 0)
                return BadRequest("此扎号没有大板");

            if (bundle.Status != SlabBundleStatus.Manufacturing)
                return BadRequest("此扎大板状态不可评等级");

            var gra = DbContext.StoneGrades
                    .Where(g => g.Id == bundleGradeQE.GradeId)
                    .SingleOrDefault();

            if (gra == null)
                return BadRequest("石材等级不存在");


            bundle.GradeId = bundleGradeQE.GradeId;

            if (bundle.Status == SlabBundleStatus.Manufacturing && (bundle.ManufacturingState == ManufacturingState.Filled || bundle.ManufacturingState == ManufacturingState.Completed))
            {
                foreach (Slab slabForDB in bundle.Slabs)
                {
                    if (slabForDB.Status == SlabBundleStatus.Discarded)
                        continue;

                    if (slabForDB.ManufacturingState != ManufacturingState.Completed && slabForDB.ManufacturingState != ManufacturingState.Filled)
                        return BadRequest("此大板的状态不允许更改其生产状态");

                    slabForDB.ManufacturingState = ManufacturingState.Completed;
                    slabForDB.GradeId = bundle.GradeId;
                }
            }
            else
                return BadRequest("此扎大板状态不允许更改其生产状态");

            bundle.ManufacturingState = ManufacturingState.Completed;
            bundle.Type = SlabType.Polished;
            bundle.LastUpdatedTime = DateTime.Now;
            wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            DbContext.SaveChanges();
            return Success();
        }

        /// <summary>
        /// 磨抛质检确认工序结束
        /// </summary>
        /// <param name="workOrderId">生产工单IdJSON数据</param>
        /// <param name="authCode">发送钉钉消息验证码</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
                return BadRequest("生产工单不存在");

            if (wo.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单的状态不允许磨抛");

            Block block = wo.MaterialRequisition.Block;

            wo.ManufacturingState = ManufacturingState.PolishingQEFinished;
            wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;

            wo.PolishedSlabOutturnPercentage = wo.AreaAfterPolishing.Value / Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value);
            wo.PolishedSlabOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.PolishedSlabOutturnPercentage.Value);

            DbContext.SaveChanges();

            // 给磨抛主管发送钉钉消息
            string polishingInfoPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("磨抛质检数据已提交，请进行磨抛确认\n荒料编号：{0} ", block.BlockNumber);
            string text = string.Format("工单 {0} 的磨抛质检数据已提交，请进行磨抛确认", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, polishingInfoPageUrl, RoleDefinition.SlabPolishingManager, authCode);
            return Success();
        }

        /// <summary>
        /// 取消生产工单
        /// </summary>
        /// <param name="cancel">销售订单JSON数据，在数据中必须提供Id</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("Cancel")]
        public async Task<IActionResult> Cancel([FromBody] WorkOrderCancelModel cancel)
        {
            if (cancel.CancelReason == null)
                return BadRequest("请输入取消订单的原因");

            var wo = DbContext.WorkOrders
                     .Where(s => s.Id == cancel.WorkOrderId)
                     .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .SingleOrDefault();

            if (wo == null)
                return BadRequest("生产工单不存在");

            if ((wo.ManufacturingState != ManufacturingState.NotStarted && wo.ManufacturingState != ManufacturingState.MaterialRequisitionSubmitted
             && wo.ManufacturingState != ManufacturingState.MaterialRequisitioned) || wo.ManufacturingState == ManufacturingState.Cancelled)
                return BadRequest("生产工单的状态不允许取消生产工单");

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

            // 给相关人员发送钉钉消息
            string workOrderCancelPageUrl = "/workOrders/info/" + wo.Id;
            string title = "";
            if (wo.MaterialRequisition != null)
            {
                if (wo.MaterialRequisition.Block != null)
                    title = string.Format("工单已取消\n荒料编号：{0} ", wo.MaterialRequisition.Block.BlockNumber);
            }
            else
                title = "工单已取消";
            string text = string.Format("工单 {0} 已取消，请查看并停止生产", wo.OrderNumber);
            await SendDingtalkLinkMessage(title, text, workOrderCancelPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.BlockManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, cancel.AuthCode);

            return Success();

        }

        /// <summary>
        /// 获取我的任务
        /// </summary>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
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
        /// 获取生产工单序号
        /// </summary>
        /// <returns>如果成功，则返回序号，失败返回相应的HTTP错误代码</returns>
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
        /// 按扎返回补胶
        /// </summary>
        /// <param name="bundleId">大板扎的Id</param>
        /// <param name="authCode">发送钉钉消息验证码</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [HttpGet()]
        [Route("BundleGoBackFilling")]
        public async Task<IActionResult> BundleGoBackFilling([FromQuery] int bundleId, [FromQuery] string authCode)
        {
            if (bundleId <= 0)
                return BadRequest("此扎大板Id不存在");

            var sb = DbContext.StoneBundles
                            .Where(b => b.Id == bundleId)
                            .Include(b => b.Block)
                            .Include(b => b.Slabs)
                            .SingleOrDefault();

            if (sb == null)
                return BadRequest("此扎大板不存在");

            if (sb.TotalSlabCount == 0)
                return BadRequest("此扎大板没有可用的大板");

            var mr = DbContext.MaterialRequisitions
                    .Where(m => m.BlockId == sb.BlockId)
                    .Include(m => m.WorkOrder)
                    .SingleOrDefault();

            if (mr == null)
                return BadRequest("领料单不存在");

            if (mr.WorkOrder == null)
                return BadRequest("生产工单不存在");

            if (mr.WorkOrder.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单状态不允许返回补胶");

            if (sb.Status != SlabBundleStatus.Manufacturing && sb.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此大板的状态不允许返回补胶");

            sb.ManufacturingState = ManufacturingState.Sawed;
            foreach (Slab slab in sb.Slabs)
            {
                if (slab.Status == SlabBundleStatus.Manufacturing && slab.ManufacturingState == ManufacturingState.Filled)
                    slab.ManufacturingState = ManufacturingState.Sawed;
            }
            DbContext.SaveChanges();

            // 给补胶主管发消息
            string workOrderPageUrl = "/workorders/info/" + mr.WorkOrder.Id;
            string title = "大板返回补胶";
            string text = string.Format("{0} 扎大板从磨抛质检处返回补胶，请完成补胶后更新补胶数据", string.Format("{0} {1}-{2}", sb.BlockNumber, sb.TotalBundleCount, sb.BundleNo));
            await SendDingtalkLinkMessage(title, text, workOrderPageUrl, RoleDefinition.FillingManager, authCode);
            return Success();
        }

        /// <summary>
        /// 按片返回补胶
        /// </summary>
        /// <param name="slabId">大板Id</param>
        /// <param name="authCode">钉钉临时授权码，如果不传，则不进行钉钉通知</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [HttpGet()]
        [Route("SlabGoBackFilling")]
        public async Task<IActionResult> SlabGoBackFilling([FromQuery] int slabId, [FromQuery] string authCode)
        {
            if (slabId <= 0)
                return BadRequest("此扎大板Id不存在");

            var slab = DbContext.Slabs
                       .Where(s => s.Id == slabId)
                       .Include(s => s.Bundle)
                       .Include(s => s.Block)
                       .SingleOrDefault();

            if (slab == null)
                return BadRequest("此大板不存在");

            var bundle = DbContext.StoneBundles
                        .Where(b => b.Id == slab.BundleId)
                        .Include(b => b.Slabs)
                        .SingleOrDefault();

            if (bundle == null)
                return BadRequest("此大板对应的扎不存在");

            if (bundle.Status != SlabBundleStatus.Manufacturing && bundle.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("大板对应的扎状态不允许返回补胶");

            if (slab.Status != SlabBundleStatus.Manufacturing && slab.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此大板的状态不允许返回补胶");

            var mr = DbContext.MaterialRequisitions
                    .Where(m => m.BlockId == slab.BlockId)
                    .Include(m => m.WorkOrder)
                    .SingleOrDefault();

            if (mr == null)
                return BadRequest("领料单不存在");

            if (mr.WorkOrder == null)
                return BadRequest("生产工单不存在");

            if (mr.WorkOrder.ManufacturingState != ManufacturingState.Filled)
                return BadRequest("此工单状态不允许返回补胶");

            slab.ManufacturingState = ManufacturingState.Sawed;
            bundle.ManufacturingState = Helper.JudgeBundleGoBackFilling(bundle);

            DbContext.SaveChanges();

            // 给补胶主管发消息
            string workOrderPageUrl = "/workorders/info/" + mr.WorkOrder.Id;
            string title = "大板返回补胶";
            string text = string.Format("{0} 扎中的序号为 {1} 的大板从磨抛质检处返回补胶，请完成补胶后更新补胶数据", string.Format("{0} {1}-{2}", bundle.BlockNumber, bundle.TotalBundleCount, bundle.BundleNo), slab.SequenceNumber);
            await SendDingtalkLinkMessage(title, text, workOrderPageUrl, RoleDefinition.FillingManager, authCode);

            return Success();
        }

        /// <summary>
        /// 报废指定的工单中的荒料
        /// </summary>
        /// <param name="workOrderId">荒料JSON数据，在数据中必须提供Id</param>
        /// <param name="discardedReason">报废原因，在数据中必须提供报废原因</param>
        /// <param name="authCode">发送钉钉消息验证码</param>
        /// <returns>如果成功，则返回标准Success数据结构，失败返回相应的HTTP错误代码</returns>
        [HttpGet]
        [Route("DiscardBlock")]
        public async Task<IActionResult> DiscardBlock([FromQuery] int workOrderId, [FromQuery] string discardedReason, [FromQuery] string authCode)
        {
            if (workOrderId <= 0)
                return BadRequest("工单Id不合法");

            discardedReason = discardedReason.Trim();
            if (string.IsNullOrEmpty(discardedReason))
                return BadRequest("荒料报废原因不能为空");

            var wo = DbContext.WorkOrders
                .Where(w => w.Id == workOrderId)
                .Include(w => w.MaterialRequisition)
                    .ThenInclude(m => m.Block)
                .SingleOrDefault();

            if (wo == null)
                return BadRequest("工单Id对应的工单不存在");

            if (!(wo.ManufacturingState == ManufacturingState.TrimmingDataSubmitted || wo.ManufacturingState == ManufacturingState.SawingDataSubmitted))
                return BadRequest("此生产工单生产状态不能进行荒料报废操作");

            var mr = wo.MaterialRequisition;

            if (mr == null)
                return BadRequest("工单没有对应的领料单");

            if (mr.Block == null)
                return BadRequest("工单对应的领料单不是荒料领料单或者领料单中的荒料不存在");

            var blo = mr.Block;

            if (blo.Status != BlockStatus.Manufacturing)
                return BadRequest("此荒料不在生产状态，不能进行报废操作");

            ManufacturingState status = wo.ManufacturingState;
            wo.ManufacturingState = ManufacturingState.Completed;
            wo.BlockDiscarded = true;
            wo.LastUpdatedTime = blo.LastUpdatedTime = DateTime.Now;
            blo.Status = BlockStatus.Discarded;
            blo.DiscardedReason = discardedReason;
            DbContext.SaveChanges();

            // 给相关人员发送钉钉消息
            string processName = (status == ManufacturingState.TrimmingDataSubmitted) ? "修边工序" : "大切工序";

            string workOrderCancelPageUrl = "/workOrders/info/" + wo.Id;
            string title = string.Format("荒料报废\n荒料编号：{0} ", blo.BlockNumber);
            string text = string.Format("工单 {0} 的荒料 {1} 在{2}报废，请查看并取消工单的后续生产计划", wo.OrderNumber, blo.BlockNumber, processName);
            await SendDingtalkLinkMessage(title, text, workOrderCancelPageUrl, new string[] {RoleDefinition.SawingManager,RoleDefinition.FillingManager
            ,RoleDefinition.SlabPolishingManager,RoleDefinition.ProductManager}, authCode);

            return Success();
        }

        /// <summary>
        /// 导入库存大板数据
        /// </summary>
        /// <param name="specFile">库存大板数据文件信息</param>
        /// <returns>导入结果，其中包含警告和错误消息</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("ImportBundleInStock")]
        [Authorize(Roles = RoleDefinition.DataOperator)]
        public async Task<IActionResult> ImportBundleInStock([FromBody] BundleSpecFileViewModel specFile)
        {
            string base64Content = specFile.FileContent;

            if (string.IsNullOrEmpty(specFile.FileName))
                return BadRequest("文件名不能为空");

            if (!specFile.FileName.EndsWith(".xlsx"))
                return Error("上传的文件必须是Excel 2007以后格式的.xlsx文件");

            if (string.IsNullOrEmpty(base64Content))
                return BadRequest("文件内容不能为空");

            byte[] fileBytes = null;
            try
            {
                fileBytes = Convert.FromBase64String(base64Content);
            }
            catch
            {
                return Error("文件内容不合法，导入失败");
            }

            List<string> dataParsingMsgs = new List<string>();
            List<string> dataUpdatingMsgs = new List<string>();
            List<StoneBundle> importBundles = new List<StoneBundle>();

            logger.LogTrace("从文件中读取大扎信息");

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
                logger.LogTrace("导入库存大扎数据 - {bundleNo}", bundleNo);

                // 大扎对应的数据库信息
                var bundleInDB = DbContext.StoneBundles
                    .Include(sb => sb.Slabs)
                    .Include(sb => sb.Category)
                    .Include(sb => sb.Grade)
                    .Where(sb => sb.BlockNumber == importBundle.BlockNumber && sb.BundleNo == importBundle.BundleNo)
                    .SingleOrDefault();

                // 数据库中不存在这扎大板，直接导入
                if (bundleInDB == null)
                {
                    // 找对应的荒料，如果找到则把荒料的状态设置为已生产完成
                    var block = DbContext.Blocks
                        .Where(b => b.BlockNumber == importBundle.BlockNumber)
                        .SingleOrDefault();
                    if (block != null && block.Status != BlockStatus.InStock && block.Status != BlockStatus.Processed)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("大扎对应的荒料在系统中已经存在，且状态不在库存，不允许导入此扎大板，扎号：{0}", bundleNo));
                        continue;
                    }

                    // 从数据库中找到石材种类，如果找不到导入文件中指定的石材种类，则不导入此扎
                    var category = getStoneCategory(importBundle.Category.Name);
                    if (category == null)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("导入的大扎的石材种类在系统中不存在，不能导入改大扎，扎号：{0}", bundleNo));
                        continue;
                    }
                    else
                    {
                        importBundle.CategoryId = category.Id;
                    }

                    // 从数据库中找到等级信息，如果码单的等级在系统中不存在，则使用“未知”等级
                    var grade = getStoneGrade(importBundle.Grade.Name);
                    if (grade == null)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的等级在系统中不存在，使用“未知”等级更新大扎，扎号：{0}", bundleNo));
                        grade = getStoneGrade("未知");
                    }
                    importBundle.GradeId = grade.Id;

                    // 从数据库中找到D99-99库存区域（截至6/30/17，D99-99作为默认的导入大扎的库存区域，如果有需要以后在模板中增加库存区域的栏并从文件中读取和导入）
                    var psa = getProductStockingArea("D", ProductStockingAreaType.AShelf, 99, 99);
                    importBundle.StockingAreaId = psa.Id;

                    importBundle.StockInOperator = dataOperatorName;
                    importBundle.StockInTime = DateTime.Now;
                    await InitializeDBRecord(importBundle);
                    DbContext.StoneBundles.Add(importBundle);

                    // 更新对应荒料的状态（如果存在荒料，将其从库存状态更新到已完成生产）
                    if (block != null && block.Status == BlockStatus.InStock)
                        block.Status = BlockStatus.Processed;

                    addInfoMsg(dataUpdatingMsgs, string.Format("成功导入大扎数据，扎号：{0}", bundleNo));
                }
                else
                {
                    // 数据库中有这扎大板，判断大扎状态决定是否更新
                    bool shouldUpdate = bundleInDB.NotVerified && bundleInDB.Slabs.Count == 0;
                    if (!shouldUpdate)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("导入的大扎数据在数据库中已经存在且是经过系统生产过程产生的大扎，不能导入数据，扎号：{0}", bundleNo));
                        continue;
                    }

                    if (bundleInDB.Status != SlabBundleStatus.InStock)
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("数据库中对应的大扎不是在库存状态，不能更新数据，扎号：{0}", bundleNo));
                        continue;
                    }

                    // 如果所有信息都一致，无需更新
                    if (bundleInDB.TotalBundleCount == importBundle.TotalBundleCount &&
                        bundleInDB.TotalSlabCount == importBundle.TotalSlabCount &&
                        bundleInDB.LengthAfterStockIn == importBundle.LengthAfterStockIn &&
                        bundleInDB.WidthAfterStockIn == importBundle.WidthAfterStockIn &&
                        bundleInDB.Category.Name == importBundle.Category.Name &&
                        bundleInDB.Grade.Name == importBundle.Grade.Name &&
                        bundleInDB.Area == importBundle.Area &&
                        bundleInDB.Thickness == importBundle.Thickness)
                    {
                        addInfoMsg(dataUpdatingMsgs, string.Format("导入的大扎数据和数据库中数据一致，无需更新，扎号：{0}", bundleNo));
                        continue;
                    }

                    // 如果大扎石材种类不一致，更新大扎的石材种类，如果找不到导入文件中指定的石材种类，则不导入此扎
                    if (bundleInDB.Category.Name != importBundle.Category.Name)
                    {
                        var category = getStoneCategory(importBundle.Category.Name);
                        if (category == null)
                        {
                            addErrorMsg(dataUpdatingMsgs, string.Format("导入的大扎的石材种类在系统中不存在，不能导入改大扎，扎号：{0}", bundleNo));
                            continue;
                        }
                        else
                        {
                            bundleInDB.CategoryId = category.Id;
                        }
                    }

                    // 如果大扎等级不一致，更新大扎的等级信息，如果码单的等级在系统中不存在，则使用“未知”等级
                    if (bundleInDB.Grade.Name != importBundle.Grade.Name)
                    {
                        var grade = getStoneGrade(importBundle.Grade.Name);
                        if (grade == null)
                        {
                            addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的等级在系统中不存在，使用“未知”等级更新大扎，扎号：{0}", bundleNo));
                            grade = getStoneGrade("未知");
                        }
                        bundleInDB.GradeId = grade.Id;
                    }

                    // 更新其他数据
                    bundleInDB.TotalBundleCount = importBundle.TotalBundleCount;
                    bundleInDB.TotalSlabCount = importBundle.TotalSlabCount;
                    bundleInDB.LengthAfterStockIn = importBundle.LengthAfterStockIn;
                    bundleInDB.WidthAfterStockIn = importBundle.WidthAfterStockIn;
                    bundleInDB.Area = importBundle.Area;
                    bundleInDB.Thickness = importBundle.Thickness;
                    bundleInDB.LastUpdatedTime = DateTime.Now;
                    addInfoMsg(dataUpdatingMsgs, string.Format("成功更新数据库中的大扎数据，扎号：{0}", bundleNo));
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
        /// 从数据库中通过石材种类名称获取石材种类信息
        /// </summary>
        /// <param name="categoryName">石材种类名称</param>
        /// <returns>StoneCategory对象，如果找不到则返回null</returns>
        StoneCategory getStoneCategory(string categoryName)
        {
            var category = DbContext.StoneCategories
                            .Where(sc => sc.Name == categoryName)
                            .SingleOrDefault();
            return category;
        }

        /// <summary>
        /// 从数据库中通过石材等级名称获取石材等级信息
        /// </summary>
        /// <param name="categoryName">石材等级名称</param>
        /// <returns>StoneGrade对象，如果找不到则返回null</returns>
        StoneGrade getStoneGrade(string gradeName)
        {
            var grade = DbContext.StoneGrades
                            .Where(g => g.Name == gradeName)
                            .SingleOrDefault();
            return grade;
        }

        /// <summary>
        /// 导入大扎磨抛质检数据
        /// </summary>
        /// <param name="specFile">大扎磨抛质检数据文件信息</param>
        /// <returns>导入结果，其中包含警告和错误消息</returns>
        [NullModelNotAllowed]
        [ValidateModelState]
        [HttpPost]
        [Route("ImportPolishingData")]
        [Authorize(Roles = RoleDefinition.DataOperator)]
        public async Task<IActionResult> ImportPolishingData([FromBody] BundleSpecFileViewModel specFile)
        {
            string base64Content = specFile.FileContent;

            if (string.IsNullOrEmpty(specFile.FileName))
                return BadRequest("文件名不能为空");

            if (!specFile.FileName.EndsWith(".xlsx"))
                return Error("上传的文件必须是Excel 2007以后格式的.xlsx文件");

            if (string.IsNullOrEmpty(base64Content))
                return BadRequest("文件内容不能为空");

            byte[] fileBytes = null;
            try
            {
                fileBytes = Convert.FromBase64String(base64Content);
            }
            catch
            {
                return Error("文件内容不合法，导入失败");
            }

            List<string> dataParsingMsgs = new List<string>();
            List<string> dataUpdatingMsgs = new List<string>();
            List<StoneBundle> importBundles = new List<StoneBundle>();

            logger.LogTrace("从文件中读取大扎信息");

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
                logger.LogTrace("导入大扎磨抛数据 - {bundleNo}", bundleNo);

                if (importBundle.Slabs.Count == 0)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("码单大扎没有任何合法的大板数据，扎号：{0}", bundleNo));
                    continue;
                }

                // 大扎对应的荒料信息
                var block = DbContext.Blocks
                    .Include(b => b.Bundles)
                        .ThenInclude(sb => sb.Slabs)
                    .Where(b => b.BlockNumber == importBundle.BlockNumber)
                    .SingleOrDefault();

                if (block == null)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("大扎对应的荒料在系统中不存在，扎号：{0}", bundleNo));
                    continue;
                }

                if (block.Status != BlockStatus.Manufacturing)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("大扎对应的荒料不是在生产状态，扎号：{0}", bundleNo));
                    continue;
                }

                // 找出数据库中对应的大扎
                StoneBundle bundleInDB = block.Bundles.ToList().Find(sb => sb.BundleNo == importBundle.BundleNo);

                if (bundleInDB == null)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("码单中的大扎在系统不存在，扎号：{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Slabs.Count == 0)
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("码单中的大扎在系统中没有任何大板数据，扎号：{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Status != SlabBundleStatus.Manufacturing || (bundleInDB.ManufacturingState != ManufacturingState.Completed && bundleInDB.ManufacturingState != ManufacturingState.Filled))
                {
                    addErrorMsg(dataUpdatingMsgs, string.Format("大扎的状态或者生产阶段不能导入磨抛数据，扎号：{0}", bundleNo));
                    continue;
                }

                if (bundleInDB.Thickness != importBundle.Thickness)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的厚度和系统中不一致，请联系管理员更新数据，扎号：{0}", bundleNo));
                }

                // 更新大扎的等级信息，如果码单的等级在系统中不存在，则使用“未知”等级
                var grade = getStoneGrade(importBundle.Grade.Name);
                if (grade == null)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的等级在系统中不存在，使用“未知”等级更新大扎，扎号：{0}", bundleNo));
                    grade = getStoneGrade("未知");
                }
                bundleInDB.Grade = grade;
                bundleInDB.GradeId = grade.Id;

                // 找出该扎大板的荒料对应的工单
                // 2017-6-27，修复BUG #632，增加工单生产工序筛选条件，只查找等待磨抛质检和等待磨抛确认的工单，避免同一颗荒料被多个工单领出来后系统出错（详见BUG #632）
                var wo = DbContext.WorkOrders
                    .Include(w => w.MaterialRequisition)
                        .ThenInclude(mr => mr.Block)
                    .Where(w => w.MaterialRequisition.Block.BlockNumber == importBundle.BlockNumber && (w.ManufacturingState == ManufacturingState.Filled || w.ManufacturingState == ManufacturingState.PolishingQEFinished))
                    .SingleOrDefault();

                foreach (Slab importSlab in importBundle.Slabs)
                {
                    Slab slab = bundleInDB.Slabs.ToList().Find(s => s.SequenceNumber == importSlab.SequenceNumber);
                    // 码单中的大板在数据库对应的大扎中不存在（可能是：1、大板从其他扎调换到此扎，2、新的大板）
                    if (slab == null)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("码单中的大板在系统的对应大扎中不存在，尝试将大板移入当前扎或导入新大板数据，扎号：{0}，大板序号：{1}", bundleNo, importSlab.SequenceNumber));
                        slab = DbContext.Slabs
                            .Include(s => s.Bundle)
                                .ThenInclude(b => b.Slabs)
                            .Include(s => s.Block)
                            .Where(s => s.SequenceNumber == importSlab.SequenceNumber && s.Block.BlockNumber == importBundle.BlockNumber)
                            .SingleOrDefault();

                        // 大板在系统中不存在，自动将新的大板导入到数据库
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
                                SawingNote = "磨抛质检阶段新加大板",
                                SequenceNumber = importSlab.SequenceNumber,
                                Status = bundleInDB.Status,
                                Thickness = bundleInDB.Thickness,
                                Type = SlabType.Polished
                            };
                            // 如果工单没有跳过补胶，则给大板添加补胶尺寸，否则让补胶尺寸留空
                            if (wo != null && wo.SkipFilling == false)
                            {
                                slab.LengthAfterFilling = slab.LengthAfterSawing;
                                slab.WidthAfterFilling = slab.WidthAfterSawing;
                            }
                            addInfoMsg(dataUpdatingMsgs, string.Format("成功将码单中的大板添加到系统中，扎号：{0}，大板序号：{1}", bundleNo, importSlab.SequenceNumber));
                        }
                        else
                        {
                            // 2017-6-27，James，修复BUG #636
                            slab.Bundle.Slabs.Remove(slab); // 从原始大扎中将要调换的大板移除
                            slab.Bundle.Area = Helper.GetAvailableSlabAreaForBundle(slab.Bundle);   // 重新计算原始大扎的面积和大板数量
                            slab.Bundle.TotalSlabCount = Helper.GetAvailableSlabCount(slab.Bundle);
                            addInfoMsg(dataUpdatingMsgs, string.Format("成功将该大板从第{2}扎移动到当前扎，当前扎号：{0}，大板序号：{1}", bundleNo, importSlab.SequenceNumber, slab.Bundle.BundleNo));
                        }

                        // 将大板添加到数据库的大扎中
                        slab.Bundle = bundleInDB;
                        bundleInDB.Slabs.Add(slab);
                    }

                    if (slab.Status != SlabBundleStatus.Manufacturing || (slab.ManufacturingState != ManufacturingState.Filled && slab.ManufacturingState != ManufacturingState.Completed))
                    {
                        addErrorMsg(dataUpdatingMsgs, string.Format("码单中大板的状态或者生产阶段不能导入磨抛数据，扎号：{0}，大板序号：{1}", bundleNo, importSlab.SequenceNumber));
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

                    addInfoMsg(dataUpdatingMsgs, string.Format("成功更新大板数据，扎号：{0}，大板序号：{1}", bundleNo, importSlab.SequenceNumber));
                }

                bundleInDB.Area = Helper.GetAvailableSlabAreaForBundle(bundleInDB);
                bundleInDB.TotalSlabCount = Helper.GetAvailableSlabCount(bundleInDB);
                bundleInDB.LastUpdatedTime = DateTime.Now;

                // 判断码单中的大扎与系统中的大扎的大板数据是否一致，如果一致则可以考虑自动完成大扎磨抛质检
                // 一致的标准：大板片数一致以及大扎中的大板序号一一对应
                bool bundleQEAutoCompletionEligible = false;    // 是否具备自动完成大扎磨抛质检的条件
                if (bundleInDB.TotalSlabCount != importBundle.TotalSlabCount)
                {
                    addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的总片数和系统中对应大扎总片数不一致，不能自动完成大扎的磨抛质检，请复核数据后手工完成大扎磨抛质检，扎号：{0}", bundleNo));
                }
                else
                {
                    List<int> slabSeqNumInDB = bundleInDB.Slabs.ToList().FindAll(s => s.Status == SlabBundleStatus.Manufacturing).Select(s => s.SequenceNumber).ToList();
                    List<int> slabSeqNumInFile = importBundle.Slabs.ToList().Select(s => s.SequenceNumber).ToList();

                    // 仅对比数据库大扎中的总片数和码单中人工填写的大扎总片数是不够的，还要两边符合条件的大板总片数数据（真实大板片数)
                    // 以下语句避免这种情况：数据库中的总片数和码单中人工填写总片数一致，但由于码单中大板数据质量问题导致大板没有被添加到内存中的大扎中
                    // 这种情况出现时，循环对比两个大扎的大板序号可能会报错（DB中大扎总片数大于内存中导入的大扎的实际大板片数导致OutOfRangeException）
                    if (slabSeqNumInDB.Count != slabSeqNumInFile.Count)
                    {
                        addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的总片数和系统中对应大扎总片数不一致，不能自动完成大扎的磨抛质检，请复核数据后手工完成大扎磨抛质检，扎号：{0}", bundleNo));
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
                                addWarningMsg(dataUpdatingMsgs, string.Format("码单中大扎的大板序号和系统中大扎的大板序号没有一一对应，不能自动完成大扎的磨抛质检，请复核数据后手工完成大扎磨抛质检，扎号：{0}", bundleNo));
                                break;
                            }
                        }
                        bundleQEAutoCompletionEligible = slabSeqNumMatched;
                    }
                }

                // 如果具备自动完成大扎磨抛质检的条件，则检查大扎下的所有大板状态，如果大板状态为报废或者Completed，则把大扎状态也置为Completed
                if (bundleQEAutoCompletionEligible)
                {
                    bool allSlabQEFinished = bundleInDB.Slabs.All(s => s.ManufacturingState == ManufacturingState.Completed || s.Status == SlabBundleStatus.Discarded);
                    if (allSlabQEFinished)
                    {
                        bundleInDB.ManufacturingState = ManufacturingState.Completed;
                        addInfoMsg(dataUpdatingMsgs, string.Format("成功完成大扎磨抛质检，扎号：{0}", bundleNo));
                    }
                }

                // 更新工单的统计数据
                if (wo != null)
                {
                    wo.AreaAfterPolishing = Helper.GetBlockManufacturingArea(block, ManufacturingState.Completed);
                    wo.PolishingQE = (await UserManager.FindByNameAsync(User.Identity.Name)).Name;
                    wo.LastUpdatedTime = DateTime.Now;
                }

                //检查该颗荒料下所有大扎状态，如果全部都为Completed，则把工单状态置为下一个工序状态，更新光板出材率并且发送通知
                bool allBundleQEFinished = block.Bundles.All(sb => sb.ManufacturingState == ManufacturingState.Completed);
                if (allBundleQEFinished)
                {
                    wo.ManufacturingState = ManufacturingState.PolishingQEFinished;

                    wo.PolishedSlabOutturnPercentage = wo.AreaAfterPolishing.Value / Helper.CalculateVolume(block.TrimmedLength.Value, block.TrimmedWidth.Value, block.TrimmedHeight.Value);
                    wo.PolishedSlabOutturnPercentage = Helper.GetThreeDecimalPlaces(wo.PolishedSlabOutturnPercentage.Value);

                    addInfoMsg(dataUpdatingMsgs, string.Format("成功完成工单的磨抛质检工序，荒料号：{0}，工单号：{1}", block.BlockNumber, wo.OrderNumber));

                    // 给磨抛主管发送钉钉消息
                    string polishingInfoPageUrl = "/workOrders/info/" + wo.Id;
                    string title = string.Format("磨抛质检数据已提交，请进行磨抛确认\n荒料编号：{0} ", block.BlockNumber);
                    string text = string.Format("工单 {0} 的磨抛质检数据已提交，请进行磨抛确认", wo.OrderNumber);
                    await SendDingtalkLinkMessage(title, text, polishingInfoPageUrl, RoleDefinition.SlabPolishingManager, specFile.AuthCode);
                }

                DbContext.SaveChanges();
            }

            return Success(new { DataUpdatingMessages = dataUpdatingMsgs, DataParsingMessages = dataParsingMsgs });
        }

        void addErrorMsg(List<string> msgs, string msg)
        {
            logger.LogError(msg);
            addMsg(msgs, string.Format("错误：{0}", msg));
        }
        void addWarningMsg(List<string> msgs, string msg)
        {
            logger.LogWarning(msg);
            addMsg(msgs, string.Format("警告：{0}", msg));
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
