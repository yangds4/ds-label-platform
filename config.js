// 数据库与服务配置（内网 Windows 服务器部署时按需修改）
module.exports = {
  server: {
    port: 3000,          // 服务端口
    host: '0.0.0.0'      // 监听所有网卡，内网其他机器可访问
  },
  mysql: {
    host: '127.0.0.1',
    port: 3306,
    user: 'root',
    password: '123456',  // 改成你的 MySQL 密码
    database: 'annotation_v2',
    charset: 'utf8mb4',
    connectionLimit: 20
  },
  // 标注员一次领取的组数（每组=1张图片）
  claimBatchSize: 50,
  // 主管名单：这些姓名在标注工作台可查看所有人已提交的标注（按标注员分组，只读）
  supervisors: ['杨东升']
};
