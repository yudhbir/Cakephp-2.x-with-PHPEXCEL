<?php

App::uses('AppController', 'Controller');

/**
 * BasicreportsController Controller
 * author: Yudhbir Singh 
 *
 * @property Contract $Contract
 * @property PaginatorComponent $Paginator
 */
class BasicreportsController extends AppController {
    
    /**
     * Components
     *
     * @var array
     */
    public $components = array('Paginator');
    public $limit=50;

    function beforeFilter() {
        parent::beforeFilter();
        /**
         * Stores array of deniable methods, without logging in.
         */
        $this->_deny = array(
            'manager' => array(
                'manager_all_reports',
                'manager_employee_creation',
                'manager_experiences_delivery',
                'manager_experiences_completion',
                'manager_download_basic_report',
            )
        );

        $this->_deny_url($this->_deny);
        $this->helpers[] = 'Custom';
    }
    public function manager_all_reports(){
        $this->layout = 'apricot';
    }
    public function manager_employee_creation(){
        $this->layout = 'apricot';
        $this->__checkIsValidEditor(array('model'=>'User','session'=>$this->_manager_data));       
        $manager = $this->Session->read('manager');
        $this->loadModel('User');
        $this->loadModel('Organization');

        $timezone = $this->Organization->field('timezone', array('id' => $this->_manager_data['organization_id']));
        $conditions=array(
            "User.group_id != " . ADMIN_GROUP_ID,"User.group_id != " . MANAGER_GROUP_ID,"User.organization_id" => $manager['User']['organization_id'],
        );
       
        if(!empty($this->request->query) && !empty($this->request->query['filter'])){            
            if(!empty($this->request->query['date_from'])){
                $date_from="str_to_date(User.created_date,'%d/%m/%Y') >= str_to_date('{$this->request->query['date_from']}','%d/%m/%Y')";
                $conditions[]= $date_from;
            }
            if(!empty($this->request->query['date_to'])){
                $date_to="str_to_date(User.created_date,'%d/%m/%Y') <= str_to_date('{$this->request->query['date_to']}','%d/%m/%Y')";
                $conditions[]= $date_to;
            }
            if(!empty($this->request->query['manager']) && !in_array('all',$this->request->query['manager'])){
                $conditions["User.createdby_id"]= $this->request->query['manager'];
            }
        
            $this->User->belongsTo = array(
                'UserApprover' => array(
                    'className' => 'User',
                    'foreignKey' => 'assigned_manager',
                    'fields' => array('name')
                ),
                'Branch' => array(
                    'className' => 'Branch',
                    'foreignKey' => 'branch_access',
                    'fields' => array('id','name')
                )
            );       
            $this->User->hasMany = array(
                'Experience' => array(                
                    'foreignKey' => 'user_id',
                    'fields' => array('id'),
                    'conditions' => array('Experience.status' =>[1,2,3,4,5,6,7]),
                )
            );
                    
            $Query = ['conditions' => $conditions,'order' => 'User.id DESC','fields'=>[
                'id','fname','lname','username','created_date', 'createdby','modified','modifiedby','UserApprover.name','Branch.id','Branch.name',
            ]];
            if(!empty($this->request->query) && !empty($this->request->query['download'])){
                $users = $this->User->find('all',$Query);
                $this->download_employee_report($users,$timezone,$this->request->query);
            }else{
                $Query['limit']=$this->limit;
                $this->Paginator->settings = $Query;
                $users =  $this->Paginator->paginate('User');
            }
            $this->set('users',$users);
        }
        $mconditions=array("User.group_id" => MANAGER_GROUP_ID,"User.organization_id" => $manager['User']['organization_id']);        
        $manager = $this->User->find('all',['conditions' => $mconditions,'order' => 'User.id DESC','fields'=>['id','username']]);
        // debug($manager);
        $this->set('manager',$manager);


    }
    public function manager_experiences_delivery(){
        $this->layout = 'apricot';
        $this->__checkIsValidEditor(array('model'=>'User','session'=>$this->_manager_data));       
        $manager = $this->Session->read('manager');
        $this->loadModel('User');  
        $this->loadModel('Organization');

        $timezone = $this->Organization->field('timezone', array('id' => $this->_manager_data['organization_id']));  

        $conditions=array(
            "Experience.organization_id" => $manager['User']['organization_id'],'Experience.status' =>[1,2,3,4,5,6,7,10]
        );
        if(!empty($this->request->query) && !empty($this->request->query['filter'])){
            if(!empty($this->request->query['date_from'])){
                $date_from="DATE_FORMAT(Experience.created,'%Y-%m-%d') >='".date('Y-m-d',strtotime(str_replace('/', '-', $this->request->query['date_from'])))."'";
                $conditions[]= $date_from;
            }
            if(!empty($this->request->query['date_to'])){
                $date_to="DATE_FORMAT(Experience.created,'%Y-%m-%d') <='".date('Y-m-d',strtotime(str_replace('/', '-', $this->request->query['date_to'])))."'";
                $conditions[]= $date_to;
            }
            if(!empty($this->request->query['manager']) && !in_array('all',$this->request->query['manager'])){
                $conditions["Experience.manager_id"]= $this->request->query['manager'];
            }
        
            $download_flag=false;
            if(!empty($this->request->query) && !empty($this->request->query['download'])){
                $download_flag=true;
            }
            $experiences=$this->experience_process_information($conditions,$download_flag);

            if(!empty($this->request->query) && !empty($this->request->query['download'])){
                $this->download_experience_report($experiences,'delivery',$timezone,$this->request->query);
            }
            $this->set('experiences',$experiences);            
        }
        $mconditions=array("User.group_id" => MANAGER_GROUP_ID,"User.organization_id" => $manager['User']['organization_id']);        
        $manager = $this->User->find('all',['conditions' => $mconditions,'order' => 'User.id DESC','fields'=>['id','username']]);
        // debug($experiences);
        $this->set('manager',$manager);
        $this->set('timezone',$timezone);
    }
    private function experience_process_information($conditions,$download_flag){
        $this->loadModel('Experience');
        $joins=array(
            array(
                'table' => 'users',
                'alias' => 'UserJoinz',
                'type' => 'INNER',
                'conditions' => array(
                    'UserJoinz.id = Experience.user_id',
                )
            ),
            array(
                'table' => 'users',
                'alias' => 'MUserJoinz',
                'type' => 'INNER',
                'conditions' => array(
                    'MUserJoinz.id = Experience.manager_id',
                )
            ),
            array(
                'table' => 'branches',
                'alias' => 'Branches',
                'type' => 'INNER',
                'conditions' => array(
                    'Branches.id = Experience.branch_id',
                )
            )
        );
        $this->Experience->virtualFields['delivery_status'] = '0';
        $this->Experience->virtualFields['completion_status'] = '0';
        $this->Experience->virtualFields['user_name'] = '0';
        $this->Experience->virtualFields['manager_name'] = '0';
        $this->Experience->virtualFields['branch_name'] = '0';
        $Query = ['conditions' => $conditions,'joins'=>$joins,'order' => 'Experience.id DESC','fields'=>[
            'id','user_id','CONCAT(UserJoinz.fname," ",UserJoinz.lname) as Experience__user_name','include_contract','include_form', 'include_policies','include_day_on_screen',
            "(CASE  
                WHEN Experience.status IN (1, 7) THEN 'Approved'
                WHEN Experience.status = 2 THEN 'In Review'
                WHEN Experience.status = 3 THEN 'Cancelled'
                WHEN Experience.status = 4 THEN 'Completed'
                WHEN Experience.status = 5 THEN 'Declined'
                WHEN Experience.status = 10 THEN 'Scheduled'
            ELSE 'Draft'
            END) AS Experience__delivery_status",
            "(CASE 
                WHEN Experience.status = 1 THEN 'Not Started'
                WHEN Experience.status = 2 THEN 'Not Started'
                WHEN Experience.status = 3 THEN 'Cancelled'
                WHEN Experience.status = 4 THEN 'Completed'
                WHEN Experience.status = 7 THEN 'In Progress'
            ELSE 'Not Started'
            END) AS Experience__completion_status",
            'created','modified','status','MUserJoinz.username as Experience__manager_name','Branches.name as Experience__branch_name',
        ]];
        if(!empty($download_flag)){
            return $this->Experience->find('all',$Query);
        }else{
            $Query['limit']=$this->limit;
            $this->Paginator->settings = $Query;
            return $experiences=$this->Paginator->paginate('Experience');
        }
    }
    public function manager_experiences_completion(){
        $this->layout = 'apricot';
        $this->__checkIsValidEditor(array('model'=>'User','session'=>$this->_manager_data));       
        $manager = $this->Session->read('manager');
        $this->loadModel('User');
        $this->loadModel('Organization');
        $timezone = $this->Organization->field('timezone', array('id' => $this->_manager_data['organization_id']));

        $conditions=["Experience.organization_id" => $manager['User']['organization_id'],'Experience.status' =>4];

        if(!empty($this->request->query) && !empty($this->request->query['filter'])){
            if(!empty($this->request->query['date_from'])){
                $date_from="DATE_FORMAT(Experience.created,'%Y-%m-%d') >='".date('Y-m-d',strtotime(str_replace('/', '-', $this->request->query['date_from'])))."'";
                $conditions[]= $date_from;
            }
            if(!empty($this->request->query['date_to'])){
                $date_to="DATE_FORMAT(Experience.created,'%Y-%m-%d') <='".date('Y-m-d',strtotime(str_replace('/', '-', $this->request->query['date_to'])))."'";
                $conditions[]= $date_to;
            }
            if(!empty($this->request->query['manager']) && !in_array('all',$this->request->query['manager'])){
                $conditions["Experience.manager_id"]= $this->request->query['manager'];
            }
        
            $download_flag=false;
            if(!empty($this->request->query) && !empty($this->request->query['download'])){
                $download_flag=true;
            }

            $experiences=$this->experience_process_information($conditions,$download_flag);

            if(!empty($this->request->query) && !empty($this->request->query['download'])){
                $this->download_experience_report($experiences,'completion',$timezone,$this->request->query);
            }

            $this->set('experiences',$experiences);           
        }
        $mconditions=array("User.group_id" => MANAGER_GROUP_ID,"User.organization_id" => $manager['User']['organization_id']);        
        $manager = $this->User->find('all',['conditions' => $mconditions,'order' => 'User.id DESC','fields'=>['id','username']]);
        // debug($manager);
        $this->set('manager',$manager);
        $this->set('timezone',$timezone);
    }
    private function download_employee_report($data,$timezone,$query){
       if(!empty($data)){
           $report_data[]=["Employee ID", "Username", "Fname", "Lname", "Creation Date", "Created by", "Last Updated", "Last Updated by" ,"Exp Count", "Branch"];
            foreach($data as $val){
                $info_data=[];               
                $info_data[]=$val['User']['id'];
                $info_data[]=$val['User']['username'];
                $info_data[]=$val['User']['fname'];
                $info_data[]=$val['User']['lname'];
                $info_data[]=$val['User']['created_date'];
                $info_data[]=$val['User']['createdby'];
                $info_data[]=date('d/m/Y H:i',$val['User']['modified']);
                $info_data[]=$val['User']['modifiedby'];
                $info_data[]=count($val['Experience']);
                $info_data[]=$val['Branch']['name'];
                $report_data[]=$info_data;
            }
            // debug($report_data);die;
            if(!empty($report_data)){
                $filename = "employee_creation_".time().".xls";
                $this->excel_setup($report_data,'employee',$filename,$timezone,$query); 
            }
       }
    }
    private function download_experience_report($data,$type,$timezone,$query){
        if(!empty($data)){
            $report_data[]=["Experience ID","User ID","Name","Branch","Documents","Forms","Checklist","Content","Delivery Status","Completion Status","Delivered on","Completed on","Delivered by"];
             foreach($data as $val){
                 $info_data=[];               
                 $info_data[]=$val['Experience']['id'];
                 $info_data[]=$val['Experience']['user_id'];
                 $info_data[]=$val['Experience']['user_name'];
                 $info_data[]=$val['Experience']['branch_name'];
                 $info_data[]=(!empty($val['Experience']['include_contract']))?'YES':"NO";
                 $info_data[]=(!empty($val['Experience']['include_form']))?'YES':"NO";
                 $info_data[]=(!empty($val['Experience']['include_policies']))?'YES':"NO";
                 $info_data[]=(!empty($val['Experience']['include_day_on_screen']))?'YES':"NO";
                 $info_data[]=$val['Experience']['delivery_status'];
                 $info_data[]=$val['Experience']['completion_status'];
                 $info_data[]=(!empty($val['Experience']['created']))?date('d/m/Y H:i',strtotime($val['Experience']['created'])):"";
                 if($type=="delivery"){
                    $info_data[]=(!empty($val['Experience']['modified']) && !empty($val['Experience']['status']) && $val['Experience']['status']=='4')?date('d/m/Y H:i',strtotime($val['Experience']['modified'])):"";
                 }else{
                    $info_data[]=(!empty($val['Experience']['modified']))?date('d/m/Y H:i',strtotime($val['Experience']['modified'])):"";
                 }
                 $info_data[]=(!empty($val['Experience']['manager_name']))?$val['Experience']['manager_name']:"";;
                 $report_data[]=$info_data;
             }
             // debug($report_data);die;
             if(!empty($report_data)){
                 $filename = "experience_".$type."_".time().".xls";
                 $this->excel_setup($report_data,$type,$filename,$timezone,$query); 
             }
        }
    }
    private function execute_download($report_data,$filename){
        $folder = WWW_ROOT.'files/';
        $full_path=$folder.$filename;
        App::uses('Folder', 'Utility');            
        $path = new Folder($folder,true,0777);                

        $file = fopen($full_path,"w");
        foreach($report_data as $val){
            fputcsv($file, $val);
        }
        $size   = filesize($full_path);
        header('Content-Description: File Transfer');
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename=' . $filename); 
        header('Content-Transfer-Encoding: binary');
        header('Connection: Keep-Alive');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');
        header('Content-Length: ' . $size);
        readfile($full_path);
        unlink($full_path);
        die(); 
    }

    public function excel_setup($report_data,$type,$filename,$timezone,$query){ 

        App::import('Thirdparty', 'PHPExcel');
        
        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties()->setCreator("Yudhbir Singh")->setLastModifiedBy("Yudhbir Singh")->setTitle("Office 2007 XLSX Test Document")
		->setSubject("Office 2007 XLSX Test Document")->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")->setKeywords("office 2007 openxml php")->setCategory("Simon Experience Basic Report");


        // Add some data
        $styleArray = array(
            'font' => array( 'bold' => true, 'color' => array('rgb' => 'FFFFFF'), 'size'  => 12, ),
            'alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY,'vertical' => PHPExcel_Style_Alignment::VERTICAL_JUSTIFY),            
            'fill' => array( 'type' => PHPExcel_Style_Fill::FILL_SOLID, 'rotation' => 90, 'startcolor' => array( 'argb' => 'FF3276b1',)),
        );  
        $Default_style = array( 'font'  => array('size'  => 12));       

        $objPHPExcel->getActiveSheet()->getStyle("A1:M1")->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->mergeCells('A1:M1');
        

        $objPHPExcel->getActiveSheet()->getStyle('A2:M2')->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->mergeCells('A2:M2');
        $r_type=(!empty($type) && $type=='employee')?'Employee By Creation Date':'Experiences By '.ucfirst($type).' Date';
        $objPHPExcel->getActiveSheet()->setCellValue('A2','Report: '.$r_type);
        

        $objPHPExcel->getActiveSheet()->getStyle('A3:M3')->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->mergeCells('A3:M3');
        $r_type_filter=(!empty($type) && $type=='employee')?'Employee Creation':'Experiences '.ucfirst($type);
        $objPHPExcel->getActiveSheet()->setCellValue('A3',$r_type_filter.' between:'.$query['date_from'].' and '.$query['date_to']);

        $objPHPExcel->getActiveSheet()->getStyle('A4:M4')->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->mergeCells('A4:M4');
        $objPHPExcel->getActiveSheet()->setCellValue('A4','Data TimeZone: '.$timezone);

        $objPHPExcel->getActiveSheet()->getStyle("A5:M5")->applyFromArray($styleArray);
        $objPHPExcel->getActiveSheet()->mergeCells('A5:M5');
        
        $objPHPExcel->getActiveSheet()->fromArray($report_data,'','A8');
        for ($i = 'A'; $i !=  $objPHPExcel->getActiveSheet()->getHighestColumn(); $i++) {
            $objPHPExcel->getActiveSheet()->getColumnDimension($i)->setAutoSize(TRUE);
        }
        // $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(1, 5, 'Php');
        // Rename worksheet
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
        $objPHPExcel->getDefaultStyle()->applyFromArray($Default_style);
        $objPHPExcel->getActiveSheet()->setTitle(substr(pathinfo($filename, PATHINFO_FILENAME),0,31));


        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$filename.'');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }

}