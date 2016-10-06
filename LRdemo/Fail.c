Fail()
{
	lr_start_transaction(lr_eval_string("Trans{rndTrans}"));
		lr_think_time(2);
	lr_end_transaction(lr_eval_string("Trans{rndTrans}"), LR_FAIL);

	lr_think_time(2);

	return 0;
}
