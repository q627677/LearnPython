#! python3
# 自动备份文件成zip文件，第一个参数是要备份的文件夹，第二个参数是备份出的文件放在哪个文件夹

import zipfile,os

def backupToZip(folder,backfolder):#定义备份函数
    os.chdir(backfolder) #更改工作目录到E盘，即备份到E盘
    folder=os.path.abspath(folder)#确保文件路径为绝对路径

    #因为之前有过备份，如果文件名重复就重新找文件名
    number=1
    while True:
        zipFilename=os.path.basename(folder)+'_'+str(number)+'.zip'
        if not os.path.exists(zipFilename):#如果不存在文件名，跳出循环
            break
        number=number+1
    
    #需要创建zip文件
    print('创建%s...'%(zipFilename))
    backupZip=zipfile.ZipFile(zipFilename,'w')
    #遍历目录
    for foldername,subfolders,filenames in os.walk(folder):
        print('添加文件夹%s...'%(foldername))
        backupZip.write(foldername)#添加文件夹到压缩文件.
        
        for filename in filenames:#(?)
            #如果文件开头是 要备份的文件夹名 和"_"，说明是此脚本的备份文件，不备份，继续循环
            #if filename.startswith(os.path.basename(folder)+'_') and filename.endswith('.zip'): 似乎此代码无用
                #continue
            backupZip.write(os.path.join(foldername,filename))#备份文件
    backupZip.close()  

    print('备份完成')

backupToZip('E:\\财务资料\\','D:\\backup')
backupToZip('E:\\财务学习\\','D:\\backup')
