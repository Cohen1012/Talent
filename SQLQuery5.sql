select * from Interview_Comments

select * from project_experience

select *from interview_Info

select * from contact_Info

insert into contact_Info (Name,UpdateTime) values ('Test1',GETDATE())

select Contact_Id from Contact_Info where ISNULL(Status,'NA') = ISNULL(ISNULL(NULL,Status),'NA') and
                                                                            UpdateTime >= ISNULL(NULL, UpdateTime) and
                                                                            UpdateTime <= ISNULL(NULL, UpdateTime) and
                                                                            ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode,Cooperation_Mode),'NA') and ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place1, Place),'NA') and ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill1, Skill),'NA')