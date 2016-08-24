# poiDemo
poi导入导出示例
使用示例：
/**
	 * 批量导入
	 * @param request
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value="importUsers", method=RequestMethod.POST)
	@ResponseBody
	@LoginVerify
	public Object importUsers(HttpServletRequest request) throws Exception {
		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
		MultipartFile multipartFile = multipartRequest.getFile("file");
		
		ResultVO<Object> result = new ResultVO<>();
		
		 NfUserVO operUser = SessionUtils.getSysOperator(request);
	    // 创建一个可读的 Excel 对象
		ReadExcelUtil excel = new ReadExcelUtil(multipartFile.getInputStream());
	    
	    // 解析 Excel 数据为 JavaBean 对象
	    List<NfUserVO> users = excel.asList(NfUserVO.class, "username", "userType", "mobile", "identifyNo", "enabled", "isSys", "email", "birthday", "sex", "hrcode");
	    Set<String> hasSame = new HashSet<String>();
	    if (users != null && users.size() > 0) {
	    	for (NfUserVO userVO : users) {
				if (StringUtil.isNullOrEmpty(userVO.getMobile()) || StringUtil.isNullOrEmpty(userVO.getIdentifyNo())
						|| StringUtil.isNullOrEmpty(userVO.getHrcode())) {
					return buildViewModel(ProcessCodeEnum.FAIL.getCode(), "手机号、身份证或工号不能为空！");
				} 
				if(hasSame.contains(userVO.getMobile())){
					return buildViewModel(ProcessCodeEnum.FAIL.getCode(), "手机号重复！");
				}else{
					hasSame.add(userVO.getMobile());
				}
				if (hasSame.contains(userVO.getIdentifyNo())) {
					return buildViewModel(ProcessCodeEnum.FAIL.getCode(), "身份证号重复！");
				} else {
					hasSame.add(userVO.getIdentifyNo());
				}
				if (hasSame.contains(userVO.getHrcode())) {
					return buildViewModel(ProcessCodeEnum.FAIL.getCode(), "工号重复！");
				} else {
					hasSame.add(userVO.getHrcode());
				}
				
				
			}
	    	
	    	result = reportService.importUsers(users, operUser);
	    } else {
	    	return buildViewModel(ProcessCodeEnum.FAIL.getCode(), "文档没有用户数据");
	    }
	    
		return buildViewModel(result.getResultCode(), result.getResultMsg());
	}

// 下载模板
@RequestMapping("/downloadResource")
	@LoginVerify
	public void downloadResource(HttpServletRequest request, HttpServletResponse response) throws Exception {
		// 获取excel模板文件的绝对路径
		ClassPathResource resource = new ClassPathResource("/templates/user.xlsx");
		try {
			File file = resource.getFile();
			if (file.exists()) 
			{
				InputStream fis = new FileInputStream(file);
				BufferedInputStream bis = null;
				BufferedOutputStream bos = null;
				response.setCharacterEncoding("utf-8");
				response.setContentType("application/vnd.ms-excel;charset=UTF-8");
				response.setHeader("Content-Disposition", "attachment;filename=" + "userExcel.xlsx");
				response.setContentLength((int) file.length());
				bis = new BufferedInputStream(fis);
				bos = new BufferedOutputStream(response.getOutputStream());
				byte[] buff = new byte[2048];
				int bytesRead;
				while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
					bos.write(buff, 0, bytesRead);
				}
				fis.close();
				bis.close();
				bos.close();
			}
		} catch (IOException e) {
			throw ProcessCodeEnum.FAIL.buildProcessException("下载失败",e);
		}
	}
	
	/**
	 * 导出拆分任务列表
	 * @param request
	 * @param response
	 * @throws Exception
	 */
	@RequestMapping("/exportSplitTaskList")
	@LoginVerify
	public void exportSplitTaskList(HttpServletRequest request, HttpServletResponse response) throws Exception {
		String reportNo = request.getParameter("reportNo");
		String splitNo = request.getParameter("splitNo");
		String taskState = request.getParameter("taskState");
		Map<String,Object> params = new HashMap<String,Object>();
		if(!StringUtil.isNullOrEmpty(reportNo)){
			params.put("reportNo", reportNo);
		}
		if(!StringUtil.isNullOrEmpty(splitNo)){
			params.put("splitNo", splitNo);
		}
		if(!StringUtil.isNullOrEmpty(taskState)){
			params.put("taskState", taskState);
		}
		Workbook excel = reportService.exportSplitTaskList(params);
		response.setCharacterEncoding("utf-8");
		response.setContentType("application/vnd.ms-excel;charset=UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + "splitTaskList.xlsx");
		OutputStream out = response.getOutputStream();
		excel.write(out);
		out.flush();
		out.close();
	}


@Override
	public Workbook exportSplitTaskList(Map<String, Object> params) throws ProcessException {
		params.put("taskType", TaskTypeEnum.SPLIT.code);
		List<ExportSplitTaskVO> exportSplitTasks = commExeSqlForPiccNfDAO.queryForList("nf_report_split.exportSplitTaskList", params);
		if (exportSplitTasks != null && exportSplitTasks.size() > 0) {
			
			String fileName = "splitTaskList.xlsx";
			String[] columnNames = { "报案号", "拆分任务号", "地址", "受损农户数量", "拆分时间", "派工时间", "任务处理人", "任务状态" };
			String[] methodNames = { "getReportNo", "getTaskNo", "getAccidentAddress", "getLossFarmerNumber", "getSplitTime", "getDispatchTime", "getHandlerName", "getTaskStateLabel" };
			
			ExcelEntity<ExportSplitTaskVO> excelEntity = new ExcelEntity<ExportSplitTaskVO>(fileName, columnNames, methodNames, exportSplitTasks);
			
			Workbook excel = null;
			try {
				excel = ExcelExporter.export2Excel(excelEntity);
				return excel;
			} catch (Exception e) {
				log.error("导出拆分任务excel异常", e);
				throw ProcessCodeEnum.FAIL.buildProcessException(e);
			}
			
		}
		return null; 
		
	}
