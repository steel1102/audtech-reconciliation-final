import { Router, type IRouter } from "express";
import healthRouter from "./health";
import reconciliationRouter from "./reconciliation";

const router: IRouter = Router();

router.use(healthRouter);
router.use(reconciliationRouter);

export default router;
