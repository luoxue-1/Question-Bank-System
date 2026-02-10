from data_manager import DataManager

class QuestionManager:
    """题目管理类，负责题目的增删改查和知识点分类"""
    
    def __init__(self):
        """
        初始化题目管理器
        """
        self.data_manager = DataManager()
    
    def get_all_questions(self, bank_id=None):
        """
        获取所有题目
        
        Args:
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 题目列表
        """
        return self.data_manager.get_questions(bank_id)
    
    def get_questions_by_topic(self, topic, bank_id=None):
        """
        根据知识点获取题目
        
        Args:
            topic: 知识点
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 题目列表
        """
        questions = self.data_manager.get_questions(bank_id)
        if not topic:
            return questions
        return [q for q in questions if q.get('topic') == topic]
    
    def add_question(self, question_data, bank_id):
        """
        添加新题目
        
        Args:
            question_data: 题目数据字典，包含content、answer、topic等字段
            bank_id: 题库ID
            
        Returns:
            int: 新题目的ID
        """
        # 验证题目数据
        if not self._validate_question(question_data):
            raise ValueError("题目数据不完整")
        
        # 添加题目
        self.data_manager.add_question(question_data, bank_id)
        return question_data['id']
    
    def update_question(self, question_id, question_data):
        """
        更新题目
        
        Args:
            question_id: 题目ID
            question_data: 更新后的题目数据
            
        Returns:
            bool: 是否更新成功
        """
        # 验证题目数据
        if not self._validate_question(question_data):
            raise ValueError("题目数据不完整")
        
        return self.data_manager.update_question(question_id, question_data)
    
    def delete_question(self, question_id):
        """
        删除题目
        
        Args:
            question_id: 题目ID
        """
        self.data_manager.delete_question(question_id)
    
    def get_question(self, question_id):
        """
        获取单个题目
        
        Args:
            question_id: 题目ID
            
        Returns:
            tuple: (题目数据字典, 题库ID)
        """
        return self.data_manager.get_question_by_id(question_id)
    
    def get_all_topics(self, bank_id=None):
        """
        获取所有知识点
        
        Args:
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 知识点列表
        """
        return self.data_manager.get_unique_topics(bank_id)
    
    def _validate_question(self, question_data):
        """
        验证题目数据
        
        Args:
            question_data: 题目数据字典
            
        Returns:
            bool: 是否验证通过
        """
        required_fields = ['content', 'answer', 'topic']
        for field in required_fields:
            if field not in question_data or not question_data[field]:
                return False
        return True
    
    def search_questions(self, keyword, bank_id=None):
        """
        根据关键词搜索题目
        
        Args:
            keyword: 搜索关键词
            bank_id: 题库ID，None表示所有题库
            
        Returns:
            list: 匹配的题目列表
        """
        questions = self.data_manager.get_questions(bank_id)
        if not keyword:
            return questions
        
        keyword = keyword.lower()
        return [q for q in questions if 
                keyword in q.get('content', '').lower() or 
                keyword in q.get('answer', '').lower() or 
                keyword in q.get('topic', '').lower()]
    
    def get_banks(self):
        """
        获取所有题库
        
        Returns:
            list: 题库列表
        """
        return self.data_manager.get_banks()
    
    def add_bank(self, bank_name):
        """
        添加新题库
        
        Args:
            bank_name: 题库名称
            
        Returns:
            int: 新题库的ID
        """
        return self.data_manager.add_bank(bank_name)
    
    def update_bank_name(self, bank_id, new_name):
        """
        更新题库名称
        
        Args:
            bank_id: 题库ID
            new_name: 新题库名称
        """
        self.data_manager.update_bank_name(bank_id, new_name)
    
    def delete_bank(self, bank_id):
        """
        删除题库
        
        Args:
            bank_id: 题库ID
        """
        self.data_manager.delete_bank(bank_id)
