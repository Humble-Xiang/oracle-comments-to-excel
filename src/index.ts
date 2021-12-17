#!/usr/bin/env node
import { program } from 'commander';
import oracledb from 'oracledb';
import CommentsWorkbook from './comments-workbook';

program
  .version('0.0.2')
  .addHelpText(
    'after',
    `

Oracle 连接依赖 node-oracledb, 使用前请确保 ORACLE_HOME 和 PATH 都已配置正确。
出现乱码问题请检查客户端 NLS_LANG 是否和数据库 NLS_LANG 保持一致。细节请参考: https://docs.oracle.com/cd/E12102_01/books/AnyInstAdm784/AnyInstAdmPreInstall18.html

Example call:
  ec2e -u yourname -p yourpwd -c '127.0.0.1:1521/orcl'
  
`
  )
  .description('一个将Oracle各个表注释导出成Excel的命令行工具')
  .requiredOption('-u, --username <username>', 'oc2e must have username')
  .requiredOption('-p, --password <password>', 'oc2e must have password')
  .requiredOption('-c, --connect-string <connect-string>', 'oc2e must have connect-string, eg: 127.0.0.1:1521/orcl')
  .action(async ({ username, password, connectString }) => {
    try {
      console.log('Connecting to Oracle...');
      await oracledb.createPool({
        user: username,
        password: password,
        connectString: connectString,
        poolAlias: username,
      });
      console.log('Connected to Oracle.');
      const connection = await oracledb.getPool(username).getConnection();
      console.log('oc2e tables...');
      await new CommentsWorkbook(connection).oc2e();
      await connection.close();
    } catch (e) {
      console.error(e);
      program.outputHelp();
    }
  });

program.parse();
